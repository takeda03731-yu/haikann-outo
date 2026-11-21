import ezdxf
from ezdxf import units
import components as cn
import constants as ct
import os
import re
from pathlib import Path
import openpyxl

# DXF新規作成
doc = ezdxf.new(dxfversion="R2010")
# doc.units = units.MM  # 単位をmmに設定
msp = doc.modelspace()


#正しいファイルパスを指定（.xlsmファイル）
current_dir = os.path.dirname(os.path.abspath(__file__))
excel_file_path = os.path.join(current_dir, "data", ct.EXCEL_FILE_NAME)

pipe_data = cn.excel_to_dict(excel_file_path)
if cn.get_excel_cell_value(excel_file_path, 'E2') == None:
    route_name = ''
else:
    route_name = cn.get_excel_cell_value(excel_file_path, 'E2')
    print(f"ルート名: {route_name}")



#pipe_dataで読み込んだ辞書をinputで指定した施工日付と、前にも材料がある場合は前の材料も足した範囲の辞書とする
data = pipe_data
filtered_data = cn.filter_by_construction_date_with_previous(data)
print(filtered_data)


haikann = [{'名称': '不断水割穿孔工φ200×φ75', '方向': 'SB'}, 
           {'名称': 'GX形ダクタイル鋳鉄片受直管φ75', '方向': 'AB'}, 
           {'名称': 'GX形ダクタイル鋳鉄片受直管φ75', '方向': 'AB'}, 
           {'名称': 'GX形ダクタイル鋳鉄片受曲管φ75 45°(V)下がり', '方向': 'AB'}, 
           {'名称': 'GX形ダクタイル鋳鉄片受直管φ75（甲切管）(V)', '方向': 'AB'}, 
           {'名称': 'GX形ダクタイル鋳鉄ソケットφ75', '方向': 'AA'}, 
           {'名称': 'GX形ダクタイル鋳鉄片受曲管φ75 45°（V）', '方向': 'BA'}, 
           {'名称': 'GX形ダクタイル鋳鉄片受直管φ75', '方向': 'BA'}
           
]

haikann2 = [{'材料No': None, '測点': 'No.0+0', '名称': None, '方向': None, '配管長': None, '水平距離': 0, '施工日付': None, '配管長2': None, '水平距離2': None, '測点2': None, '継続元': True},
            {'材料No': '17', '測点': 'No.10+200', '名称': 'メカニカルソケット', '方向': 'SA', '配管長': 200, '水平距離': 200, '施工日付': 'R6.10.15', '配管長2': None, '水平距離2': None, '測点2': None},
            {'材料No': '10', '測点': 'No.10+200', '名称': 'EF直管', '方向': 'BA', '配管長': 5050, '水平距離': 5050, '施工日付': 'R6.10.15', '配管長2': None, '水平距離2': None, '測点2': None},
            {'材料No': '1', '測点': 'No.10+200', '名称': 'EFベンド22.5°(V)（上がり）', '方向': 'AB', '配管長': 150, '水平距離': 150, '施工日付': 'R6.10.11', '配管長2': 150, '水平距離2': 130, '測点2': 'No.10+350'}, 
            {'材料No': '2', '測点': 'No.10+200', '名称': 'EF直管（甲切管）(V)', '方向': 'AB', '配管長': 1000, '水平距離': 924, '施工日付': 'R6.10.15', '配管長2': None, '水平距離2': None, '測点2': None},
            {'材料No': '3', '測点': 'No.10+200', '名称': 'EFベンド22.5°（V）', '方向': 'BA', '配管長': 150, '水平距離': 130, '施工日付': 'R6.10.15', '配管長2': 150, '水平距離2': 150, '測点2': 'No.10+350'},
            {'材料No': '4', '測点': 'No.10+200', '名称': 'EF直管', '方向': 'BA', '配管長': 5050, '水平距離': 5050, '施工日付': 'R6.10.15', '配管長2': None, '水平距離2': None, '測点2': None},
            {'材料No': '5', '測点': 'No.10+200', '名称': 'EF直管', '方向': 'BA', '配管長': 5050, '水平距離': 5050, '施工日付': 'R6.10.15', '配管長2': None, '水平距離2': None, '測点2': None},
            {'材料No': '6', '測点': 'No.10+200', '名称': 'EF直管', '方向': 'BA', '配管長': 5050, '水平距離': 5050, '施工日付': 'R6.10.15', '配管長2': None, '水平距離2': None, '測点2': None},
            {'材料No': '7', '測点': 'No.10+200', '名称': 'EF直管', '方向': 'BA', '配管長': 5050, '水平距離': 5050, '施工日付': 'R6.10.15', '配管長2': None, '水平距離2': None, '測点2': None},
            {'材料No': '8', '測点': 'No.10+200', '名称': 'EF(F付)チーズ両挿（消火栓）', '方向': 'BB', '配管長': 150, '水平距離': 150, '施工日付': 'R6.10.15', '配管長2': 150, '水平距離2': 150, '測点2': 'No.10+350'},
            {'材料No': '9', '測点': 'No.10+200', '名称': 'EFソケット', '方向': 'AA', '配管長': 200, '水平距離': 200, '施工日付': 'R6.10.15', '配管長2': 200, '水平距離2': 200, '測点2': 'No.10+5000'}
]
#位置のオフセット設定
x_count = -160000000  # SF方向の図形カウンタ
y_count = 0  # FS方向の図形カウンタ
y_offsetcount = 0
today_count = 0
today_xbase = 0
flat_sum = 0
haikann_sum = 0


# 略図の描画開始位置: '継続元'の材料（前の日の最後の材料）。それより前は'参照のみ'で描かない。
draw_start = next((i for i, d in enumerate(filtered_data) if d.get('継続元')), 0)

if filtered_data[draw_start]['材料No'] is None:
    base_x = 0 + x_count
    base_y = 0 + y_count
    msp.add_text(f"{route_name} {filtered_data[draw_start]['測点']}", dxfattribs={
    'insert': (base_x -50000000, base_y + 20000000),
    'height': 6000000,
    })
else:
    first_item = filtered_data[draw_start]
    cn.first_sokutenn(msp, first_item, route_name, offset_x=x_count, offset_y=y_count)

for i, item in enumerate(filtered_data):
    if i < draw_start:
        continue  # 参照のみの材料は描画しない（prev1〜3の判定にはfiltered_data全体を使う）
    prev1 = filtered_data[i-1] if i >= 1 else None  # 1個前
    prev2 = filtered_data[i-2] if i >= 2 else None  # 2個前
    prev3 = filtered_data[i-3] if i >= 3 else None  # 3個前

    name = item.get('名称', '')  # 安全な取得方法
    
    # 修正: Noneを除外して名称リストを作成
    names_prev = [p.get('名称', '') for p in (prev1, prev2, prev3) if p and p.get('名称') is not None]
    
    # 例：前1〜3個のどれかが「45°」かつ「下がり」を含むか
    down_90 = any(('90' in pn) and ('下がり' in pn or pn.endswith('下がり')) for pn in names_prev)
    down_45 = any(('45' in pn) and ('下がり' in pn or pn.endswith('下がり')) for pn in names_prev)
    down_22 = any(('22' in pn) and ('下がり' in pn or pn.endswith('下がり')) for pn in names_prev)
    down_11 = any(('11' in pn) and ('下がり' in pn or pn.endswith('下がり')) for pn in names_prev)
    up_90 = any(('90' in pn) and ('上がり' in pn or pn.endswith('上がり')) for pn in names_prev)
    up_45 = any(('45' in pn) and ('上がり' in pn or pn.endswith('上がり')) for pn in names_prev)
    up_22 = any(('22' in pn) and ('上がり' in pn or pn.endswith('上がり')) for pn in names_prev)
    up_11 = any(('11' in pn) and ('上がり' in pn or pn.endswith('上がり')) for pn in names_prev)

    # Noneチェック（既存のコードはそのまま）
    if name is None or name == '':
        continue  # 名称がNoneまたは空文字の場合はスキップ
    elif (
        '不断水割' in name
        and '副弁付' in name
    ):
        if item['方向'] == 'SF':
            cn.choose_zairyou_no(msp, item, offset_x=x_count + 19500000, offset_y=y_count+15000000)
            cn.under_pressure_drilling_work_ben_sf(msp,offset_x=x_count, offset_y=y_count)
            cn.name_sokutenn(msp, item, route_name, offset_x=x_count - 55000000, offset_y=y_count + 8000000 + 13500000 * y_offsetcount)
            first_x = x_count
            x_count += 39000000
            y_offsetcount += 1
            today_count += 1
            second_x = x_count
            flat_sum += item['水平距離']
            haikann_sum += item['配管長']
            msp.add_line((first_x, y_count - y_count - 45000000), (first_x, y_count - y_count - 61000000))
            msp.add_line((second_x, y_count - y_count - 45000000), (second_x, y_count - y_count - 61000000))
            msp.add_line((first_x, y_count - y_count - 61000000), (second_x, y_count - y_count - 61000000))
            cn.flat_line1(msp, item, offset_x=x_count-19500000, offset_y=y_count - y_count - 61000000)
        elif item['方向'] == 'FS':
            cn.choose_zairyou_no(msp, item, offset_x=x_count + 19500000, offset_y=y_count+15000000)
            cn.under_pressure_drilling_work_ben_fs(msp,offset_x=x_count, offset_y=y_count)
            cn.name_sokutenn(msp, item, route_name, offset_x=x_count - 55000000, offset_y=y_count + 8000000 + 13500000 * y_offsetcount)
            first_x = x_count
            x_count += 39000000
            y_offsetcount += 1
            today_count += 1
            second_x = x_count
            flat_sum += item['水平距離']
            haikann_sum += item['配管長']
            msp.add_line((first_x, y_count - y_count - 45000000), (first_x, y_count - y_count - 61000000))
            msp.add_line((second_x, y_count - y_count - 45000000), (second_x, y_count - y_count - 61000000))
            msp.add_line((first_x, y_count - y_count - 61000000), (second_x, y_count - y_count - 61000000))
            cn.flat_line1(msp, item, offset_x=x_count-19500000, offset_y=y_count - y_count - 61000000)
    elif (
        '不断水割' in name
    ):
        if item['方向'] == 'SF':
            cn.choose_zairyou_no(msp, item, offset_x=x_count, offset_y=y_count+17000000)
            cn.under_pressure_drilling_work_sf(msp,offset_x=x_count, offset_y=y_count)
            cn.name_sokutenn(msp, item, route_name, offset_x=x_count - 55000000, offset_y=y_count + 13500000 * y_offsetcount)
            first_x = x_count
            x_count += 16000000
            y_offsetcount += 1
            today_count += 1
            second_x = x_count
            flat_sum += item['水平距離']
            haikann_sum += item['配管長']
            msp.add_line((first_x, y_count - y_count - 45000000), (first_x, y_count - y_count - 61000000))
            msp.add_line((second_x, y_count - y_count - 45000000), (second_x, y_count - y_count - 61000000))
            msp.add_line((first_x, y_count - y_count - 61000000), (second_x, y_count - y_count - 61000000))
            cn.flat_line1(msp, item, offset_x=x_count-8000000, offset_y=y_count - y_count - 61000000)
            if item['水平距離'] == item['配管長']:
                pass
            else:
                cn.haikann_line1(msp, item, offset_x=x_count-19500000, offset_y=y_count - y_count - 54500000)
        elif item['方向'] == 'FE':
            cn.choose_zairyou_no(msp, item, offset_x=x_count + 15000000, offset_y=y_count+17000000)
            cn.under_pressure_drilling_work_fs(msp,offset_x=x_count, offset_y=y_count)
            cn.name_sokutenn(msp, item, route_name, offset_x=x_count - 55000000, offset_y=y_count + 13500000 * y_offsetcount)
            first_x = x_count
            x_count += 15000000
            y_offsetcount += 1
            today_count += 1
            second_x = x_count
            flat_sum += item['水平距離']
            haikann_sum += item['配管長']
            msp.add_line((first_x, y_count - y_count - 45000000), (first_x, y_count - y_count - 61000000))
            msp.add_line((second_x, y_count - y_count - 45000000), (second_x, y_count - y_count - 61000000))
            msp.add_line((first_x, y_count - y_count - 61000000), (second_x, y_count - y_count - 61000000))
            cn.flat_line1(msp, item, offset_x=x_count-8000000, offset_y=y_count - y_count - 61000000)
            if item['水平距離'] == item['配管長']:
                pass
            else:
                cn.haikann_line1(msp, item, offset_x=x_count-19500000, offset_y=y_count - y_count - 54500000)

    elif (
        '既設' in name
    ):
        if item['方向'] == 'SB':
            cn.choose_zairyou_no(msp, item, offset_x=x_count + 4000000, offset_y=y_count+15000000)
            cn.start_sb(msp,offset_x=x_count, offset_y=y_count)
            cn.name_sokutenn(msp, item, route_name, offset_x=x_count - 55000000, offset_y=y_count + 8000000 + 13500000 * y_offsetcount)
            first_x = x_count
            x_count += 8000000
            y_offsetcount += 1
            today_count += 1
            today_xbase = 8000000
            second_x = x_count

    elif (
        'GX形ダクタイル鋳鉄' in name
        and '（甲切管）' in name
        and '(V)' in name
        and down_45
    ):
        if item['方向'] == 'BA':
            cn.gx_vb_tyokkann_kou_45_down_ba(msp, item, offset_x=x_count, offset_y=y_count)
            x_count += 10250000 + (item['水平距離'] * 10000 / 2)
            y_count += -10250000 - (item['水平距離'] * 10000 / 2)
        elif item['方向'] == 'AB':
            cn.gx_vb_tyokkann_kou_45_down_ab(msp, item, offset_x=x_count, offset_y=y_count)
            x_count += 10250000 + (item['水平距離'] * 10000 / 2)
            y_count += -10250000 - (item['水平距離'] * 10000 / 2)            
    elif (
        'GX形ダクタイル鋳鉄' in name
        and '（甲切管）' in name
        and '(V)' in name
        and up_45
    ):
        if item['方向'] == 'BA':
            cn.gx_vb_tyokkann_kou_45_up_ba(msp, item, offset_x=x_count, offset_y=y_count)
            x_count += 10250000 + (item['水平距離'] * 10000 / 2)
            y_count += 10250000 + (item['水平距離'] * 10000 / 2)
        elif item['方向'] == 'AB':
            cn.gx_vb_tyokkann_kou_45_up_ab(msp, item, offset_x=x_count, offset_y=y_count)
            x_count += 10250000 + (item['水平距離'] * 10000 / 2)
            y_count += 10250000 + (item['水平距離'] * 10000 / 2)
    elif (
        'GX形ダクタイル鋳鉄' in name
        and '（甲切管）' in name
        and '(V)' in name
        and down_22
    ):
        if item['方向'] == 'BA':
            cn.gx_vb_tyokkann_kou_22_down_ba(msp, item, offset_x=x_count, offset_y=y_count)
            x_count += 13400000 + item['水平距離'] * 10000 / 2
            y_count += - 5550000 - ((item['配管長'] * 10000 / 2) ** 2 - (item['水平距離'] * 10000 / 2) ** 2) ** 0.5
        elif item['方向'] == 'AB':
            cn.gx_vb_tyokkann_kou_22_down_ab(msp, item, offset_x=x_count, offset_y=y_count)
            x_count += 13400000 + item['水平距離'] * 10000 / 2
            y_count += - 5550000 - ((item['配管長'] * 10000 / 2) ** 2 - (item['水平距離'] * 10000 / 2) ** 2) ** 0.5
    elif (
        'GX形ダクタイル鋳鉄' in name
        and '（甲切管）' in name
        and '(V)' in name
        and up_22
    ):
        if item['方向'] == 'BA':
            cn.gx_vb_tyokkann_kou_22_up_ba(msp, item, offset_x=x_count, offset_y=y_count)
            x_count += 13400000 + item['水平距離'] * 10000 / 2
            y_count += + 5550000 + ((item['配管長'] * 10000 / 2) ** 2 - (item['水平距離'] * 10000 / 2) ** 2) ** 0.5
        elif item['方向'] == 'AB':
            cn.gx_vb_tyokkann_kou_22_up_ab(msp, item, offset_x=x_count, offset_y=y_count)
            x_count += 13400000 + item['水平距離'] * 10000 / 2
            y_count += + 5550000 + ((item['配管長'] * 10000 / 2) ** 2 - (item['水平距離'] * 10000 / 2) ** 2) ** 0.5
    elif (
        'GX形ダクタイル鋳鉄' in name
        and '（甲切管）' in name
        and '(V)' in name
        and down_11
    ):
        if item['方向'] == 'BA':
            cn.gx_vb_tyokkann_kou_11_down_ba(msp, item, offset_x=x_count, offset_y=y_count)
            x_count += 14220000 + item['水平距離'] * 10000 / 2
            y_count += - 2830000 - ((item['配管長'] * 10000 / 2) ** 2 - (item['水平距離'] * 10000 / 2) ** 2) ** 0.5
        elif item['方向'] == 'AB':
            cn.gx_vb_tyokkann_kou_11_down_ab(msp, item, offset_x=x_count, offset_y=y_count)
            x_count += 14220000 + item['水平距離'] * 10000 / 2
            y_count += - 2830000 - ((item['配管長'] * 10000 / 2) ** 2 - (item['水平距離'] * 10000 / 2) ** 2) ** 0.5
    elif (
        'GX形ダクタイル鋳鉄' in name
        and '（甲切管）' in name
        and '(V)' in name
        and up_11
    ):
        if item['方向'] == 'BA':
            cn.gx_vb_tyokkann_kou_11_up_ba(msp, item, offset_x=x_count, offset_y=y_count)
            x_count += 14220000 + item['水平距離'] * 10000 / 2
            y_count += 2830000 + ((item['配管長'] * 10000 / 2) ** 2 - (item['水平距離'] * 10000 / 2) ** 2) ** 0.5
        elif item['方向'] == 'AB':
            cn.gx_vb_tyokkann_kou_11_up_ab(msp, item, offset_x=x_count, offset_y=y_count)
            x_count += 14220000 + item['水平距離'] * 10000 / 2
            y_count += 2830000 + ((item['配管長'] * 10000 / 2) ** 2 - (item['水平距離'] * 10000 / 2) ** 2) ** 0.5

    elif (
        'GX形' in name
        and '継ぎ輪' in name
    ):
        if item['方向'] == 'AA':
            if today_count == 0:
                cn.gx_tugiwa_aa(msp, offset_x=0, offset_y=0)
                today_count += 1
                x_count += 31000000
                today_xbase += 31000000     
            else:
                cn.choose_zairyou_no(msp, item, offset_x=x_count + 15500000, offset_y=y_count + 15000000)
                cn.name(msp, item, offset_x=x_count - 70000000, offset_y=y_count - y_count + 13500000 * y_offsetcount)
                cn.gx_tugiwa_aa(msp, offset_x=x_count, offset_y=y_count)
                first_x = x_count
                x_count += 31000000
                y_offsetcount += 1
                today_count += 1
                second_x = x_count
                flat_sum += (item.get('水平距離', 0) or 0) + (item.get('水平距離2', 0) or 0)
                haikann_sum += (item.get('配管長', 0) or 0) + (item.get('配管長2', 0) or 0)
                msp.add_line((first_x, y_count - y_count - 45000000), (first_x, y_count - y_count - 61000000))
                msp.add_line((second_x, y_count - y_count - 45000000), (second_x, y_count - y_count - 61000000))
                msp.add_line((first_x, y_count - y_count - 61000000), (second_x, y_count - y_count - 61000000))
                flat_line1_value = item.get('水平距離', 0) or 0
                flat_line2_value = item.get('水平距離2', 0) or 0
                goukei = flat_line1_value + flat_line2_value
                msp.add_text(str(goukei), dxfattribs={"height": 6000000}).set_placement(
                    (x_count - 20000000, y_count - y_count - 61000000)
                )

    elif (
        'GX形' in name
        and '短管' in name
        and '(V)' in name
        and down_45
    ):
        if item['方向'] == 'AA':
            if today_count == 0:
                cn.gx_ryouuke_tannkann_down_45_aa(msp, offset_x=x_count, offset_y=y_count)
                today_count += 1
                x_count += 20500000
                y_count -= 20500000
                today_xbase += 20500000          
            else:
                cn.choose_zairyou_no(msp, item, offset_x=x_count + 10250000, offset_y=y_count + 15000000)
                cn.name(msp, item, offset_x=x_count - 70000000, offset_y=y_count - y_count + 13500000 * y_offsetcount)
                cn.gx_ryouuke_tannkann_down_45_aa(msp, offset_x=x_count, offset_y=y_count)
                first_x = x_count
                x_count += 20500000
                y_count -= 20500000
                y_offsetcount += 1
                today_count += 1
                second_x = x_count
                flat_sum += (item.get('水平距離', 0) or 0) + (item.get('水平距離2', 0) or 0)
                haikann_sum += (item.get('配管長', 0) or 0) + (item.get('配管長2', 0) or 0)
                msp.add_line((first_x, y_count - y_count - 45000000), (first_x, y_count - y_count - 61000000))
                msp.add_line((second_x, y_count - y_count - 45000000), (second_x, y_count - y_count - 61000000))
                msp.add_line((first_x, y_count - y_count - 61000000), (second_x, y_count - y_count - 61000000))
                flat_line1_value = item.get('水平距離', 0) or 0
                flat_line2_value = item.get('水平距離2', 0) or 0
                goukei = flat_line1_value + flat_line2_value
                msp.add_text(str(goukei), dxfattribs={"height": 6000000}).set_placement(
                    (x_count - 14750000, y_count - y_count - 61000000)
                )

    elif (
        'GX形' in name
        and '短管' in name
        and '(V)' in name
        and up_45
    ):
        if item['方向'] == 'AA':
            if today_count == 0:
                cn.gx_ryouuke_tannkann_up_45_aa(msp, offset_x=x_count, offset_y=y_count)
                today_count += 1
                x_count += 20500000
                y_count += 20500000
                today_xbase += 20500000          
            else:
                cn.choose_zairyou_no(msp, item, offset_x=x_count + 2250000, offset_y=y_count + 20000000)
                cn.name(msp, item, offset_x=x_count - 100000000, offset_y=y_count - y_count + 13500000 * y_offsetcount)
                cn.gx_ryouuke_tannkann_up_45_aa(msp, offset_x=x_count, offset_y=y_count)
                first_x = x_count
                x_count += 20500000
                y_count += 20500000
                y_offsetcount += 1
                today_count += 1
                second_x = x_count
                flat_sum += (item.get('水平距離', 0) or 0) + (item.get('水平距離2', 0) or 0)
                haikann_sum += (item.get('配管長', 0) or 0) + (item.get('配管長2', 0) or 0)
                msp.add_line((first_x, y_count - y_count - 45000000), (first_x, y_count - y_count - 61000000))
                msp.add_line((second_x, y_count - y_count - 45000000), (second_x, y_count - y_count - 61000000))
                msp.add_line((first_x, y_count - y_count - 61000000), (second_x, y_count - y_count - 61000000))
                flat_line1_value = item.get('水平距離', 0) or 0
                flat_line2_value = item.get('水平距離2', 0) or 0
                goukei = flat_line1_value + flat_line2_value
                msp.add_text(str(goukei), dxfattribs={"height": 6000000}).set_placement(
                    (x_count - 14750000, y_count - y_count - 61000000)
                )

    elif (
        'GX形' in name
        and '短管' in name
    ):
        if item['方向'] == 'AA':
            if today_count == 0:
                cn.gx_ryouuke_tannkann_aa(msp, offset_x=0, offset_y=0)
                today_count += 1
                x_count += 29000000
                today_xbase += 29000000     
            else:
                cn.choose_zairyou_no(msp, item, offset_x=x_count + 14500000, offset_y=y_count + 15000000)
                cn.name(msp, item, offset_x=x_count - 70000000, offset_y=y_count - y_count + 13500000 * y_offsetcount)
                cn.gx_ryouuke_tannkann_aa(msp, offset_x=x_count, offset_y=y_count)
                first_x = x_count
                x_count += 29000000
                y_offsetcount += 1
                today_count += 1
                second_x = x_count
                flat_sum += (item.get('水平距離', 0) or 0) + (item.get('水平距離2', 0) or 0)
                haikann_sum += (item.get('配管長', 0) or 0) + (item.get('配管長2', 0) or 0)
                msp.add_line((first_x, y_count - y_count - 45000000), (first_x, y_count - y_count - 61000000))
                msp.add_line((second_x, y_count - y_count - 45000000), (second_x, y_count - y_count - 61000000))
                msp.add_line((first_x, y_count - y_count - 61000000), (second_x, y_count - y_count - 61000000))
                flat_line1_value = item.get('水平距離', 0) or 0
                flat_line2_value = item.get('水平距離2', 0) or 0
                goukei = flat_line1_value + flat_line2_value
                msp.add_text(str(goukei), dxfattribs={"height": 6000000}).set_placement(
                    (x_count - 19000000, y_count - y_count - 61000000)
                )

        elif item['方向'] == 'FA':
            if today_count == 0:
                cn.gx_tankan_FA(msp, offset_x=x_count, offset_y=y_count)
                x_count += 24500000
                today_count += 1
                today_xbase += 24500000
            else:
                cn.choose_zairyou_no(msp, item, offset_x=x_count + 12250000, offset_y=y_count + 15000000)
                cn.name(msp, item, offset_x=x_count - 90000000, offset_y=y_count - y_count + 13500000 * y_offsetcount)
                cn.gx_tankan_FA(msp, offset_x=x_count, offset_y=y_count)
                first_x = x_count
                x_count += 24500000
                y_offsetcount += 1
                today_count += 1
                second_x = x_count
                flat_sum += (item.get('水平距離', 0) or 0) + (item.get('水平距離2', 0) or 0)
                haikann_sum += (item.get('配管長', 0) or 0) + (item.get('配管長2', 0) or 0)
                msp.add_line((first_x, y_count - y_count - 45000000), (first_x, y_count - y_count - 61000000))
                msp.add_line((second_x, y_count - y_count - 45000000), (second_x, y_count - y_count - 61000000))
                msp.add_line((first_x, y_count - y_count - 61000000), (second_x, y_count - y_count - 61000000))
                cn.flat_line1(msp, item, offset_x=x_count - 11200000, offset_y=y_count - y_count - 61000000)

        elif item['方向'] == 'AF':
            if today_count == 0:
                cn.gx_tankan_AF(msp, offset_x=x_count, offset_y=y_count)
                x_count += 25500000
                today_count += 1
                today_xbase += 26500000
            else:
                cn.choose_zairyou_no(msp, item, offset_x=x_count + 13250000, offset_y=y_count + 15000000)
                cn.name(msp, item, offset_x=x_count - 90000000, offset_y=y_count - y_count + 13500000 * y_offsetcount)
                cn.gx_tankan_AF(msp, offset_x=x_count, offset_y=y_count)
                first_x = x_count
                x_count += 25500000
                y_offsetcount += 1
                today_count += 1
                second_x = x_count
                flat_sum += (item.get('水平距離', 0) or 0) + (item.get('水平距離2', 0) or 0)
                haikann_sum += (item.get('配管長', 0) or 0) + (item.get('配管長2', 0) or 0)
                msp.add_line((first_x, y_count - y_count - 45000000), (first_x, y_count - y_count - 61000000))
                msp.add_line((second_x, y_count - y_count - 45000000), (second_x, y_count - y_count - 61000000))
                msp.add_line((first_x, y_count - y_count - 61000000), (second_x, y_count - y_count - 61000000))
                cn.flat_line1(msp, item, offset_x=x_count - 11200000, offset_y=y_count - y_count - 61000000)

    elif (
        '直管' in name
        and '（乙切管）' in name
        and '(V)' in name
        and down_90
    ):
        if item['方向'] == 'BB':
            if today_count == 0:
                cn.vb_tyokkann_otu_90_down_bb(msp, item, offset_x=x_count, offset_y=y_count)
                y_count += - (item['配管長'] * 10000)
                msp.add_text("乙" + item['配管長'], dxfattribs={
                'insert': (x_count - 24000000, y_count - 3000000 - (item['配管長'] * 10000 / 2)),
                'height': 6000000,
                })
                today_count += 1
            else:
                cn.choose_zairyou_no(msp, item, offset_x=x_count + (item['配管長'] * 10000), offset_y=y_count + 5000000 + (item['配管長'] * 10000 / 2))                
                cn.vb_tyokkann_otu_90_down_bb(msp, item, offset_x=x_count, offset_y=y_count)
                msp.add_text(f"乙{item['配管長']}", dxfattribs={
                'insert': (x_count - 24000000, y_count - 3000000 - (item['配管長'] * 10000 / 2)),
                'height': 6000000,
                })
                y_count += - (item['配管長'] * 10000)
                today_count += 1
                flat_sum += (item.get('水平距離', 0) or 0) + (item.get('水平距離2', 0) or 0)
                haikann_sum += (item.get('配管長', 0) or 0) + (item.get('配管長2', 0) or 0)
    elif (
        '直管' in name
        and '（乙切管）' in name
        and '(V)' in name
        and up_90
    ):
        if item['方向'] == 'BB':
            if today_count == 0:
                cn.vb_tyokkann_otu_90_up_bb(msp, item, offset_x=x_count, offset_y=y_count)
                y_count += (item['配管長'] * 10000)
                msp.add_text("乙" + item['配管長'], dxfattribs={
                'insert': (x_count + 5000000, y_count - (item['配管長'] * 10000)),
                'height': 6000000,
                })
                today_count += 1
            else:
                cn.choose_zairyou_no(msp, item, offset_x=x_count + 2000000 + (item['配管長'] * 10000), offset_y=y_count -2000000 + (item['配管長'] * 10000))
                cn.vb_tyokkann_otu_90_up_bb(msp, item, offset_x=x_count, offset_y=y_count)
                msp.add_text(f"乙{item['配管長']}", dxfattribs={
                'insert': (x_count + 5000000, y_count - (item['配管長'] * 10000 / 2)),
                'height': 6000000,
                })
                y_count += (item['配管長'] * 10000)
                today_count += 1
                flat_sum += (item.get('水平距離', 0) or 0) + (item.get('水平距離2', 0) or 0)
                haikann_sum += (item.get('配管長', 0) or 0) + (item.get('配管長2', 0) or 0)
    elif (
        '直管' in name
        and '（乙切管）' in name
        and '(V)' in name
        and down_45
    ):
        if item['方向'] == 'BB':
            if today_count == 0:
                cn.vb_tyokkann_otu_45_down_bb(msp, item, offset_x=x_count, offset_y=y_count)
                x_count += (item['水平距離'] * 10000)
                y_count += - (item['水平距離'] * 10000)
                msp.add_text(f"乙{item['配管長']}", dxfattribs={
                'insert': (x_count - 20000000, y_count - (item['配管長'] * 10000)),
                'height': 6000000,
                })
                today_xbase += (item['水平距離'] * 10000)
                today_count += 1
            else:
                cn.choose_zairyou_no(msp, item, offset_x=x_count +4500000 + (item['水平距離'] * 10000 / 2), offset_y=y_count + 5000000 + (item['配管長'] * 10000))                
                cn.vb_tyokkann_otu_45_down_bb(msp, item, offset_x=x_count, offset_y=y_count)
                msp.add_text(f"乙{item['配管長']}", dxfattribs={
                'insert': (x_count - 20000000, y_count - (item['配管長'] * 10000)),
                'height': 6000000,
                })
                first_x = x_count
                x_count +=  (item['水平距離'] * 10000)
                y_count += - (item['水平距離'] * 10000)
                today_count += 1
                second_x = x_count
                flat_sum += (item.get('水平距離', 0) or 0) + (item.get('水平距離2', 0) or 0)
                haikann_sum += (item.get('配管長', 0) or 0) + (item.get('配管長2', 0) or 0)
                msp.add_line((first_x, y_count - y_count - 45000000), (first_x, y_count - y_count - 61000000))
                msp.add_line((second_x, y_count - y_count - 45000000), (second_x, y_count - y_count - 61000000))
                msp.add_line((first_x, y_count - y_count - 61000000), (second_x, y_count - y_count - 61000000))
                cn.flat_line1(msp, item, offset_x=x_count - (item['水平距離'] * 10000 / 2), offset_y=y_count - y_count - 48000000)
                if item['水平距離'] == item['配管長']:
                    pass
                else:
                    cn.haikann_line1(msp, item, offset_x=x_count -9200000 - (item['水平距離'] * 10000 / 2), offset_y=y_count - y_count - 42000000)                   
    elif (
        '直管' in name
        and '（乙切管）' in name
        and '(V)' in name
        and up_45
    ):
        if item['方向'] == 'BB':
            if today_count == 0:
                cn.vb_tyokkann_otu_45_up_bb(msp, item, offset_x=x_count, offset_y=y_count)
                x_count += (item['水平距離'] * 10000)
                y_count += (item['水平距離'] * 10000)
                msp.add_text(f"乙{item['配管長']}", dxfattribs={
                'insert': (x_count, y_count - 3000000 - (item['配管長'] * 10000)),
                'height': 6000000,
                })
                today_xbase += (item['水平距離'] * 10000)
                today_count += 1
            else:
                cn.choose_zairyou_no(msp, item, offset_x=x_count -7000000 + (item['水平距離'] * 10000 / 2), offset_y=y_count + 15000000 + (item['配管長'] * 10000))                
                cn.vb_tyokkann_otu_45_up_bb(msp, item, offset_x=x_count, offset_y=y_count)
                msp.add_text(f"乙{item['配管長']}", dxfattribs={
                'insert': (x_count, y_count - 3000000- (item['配管長'] * 10000)),
                'height': 6000000,
                })
                first_x = x_count
                x_count +=  (item['水平距離'] * 10000)
                y_count +=  (item['水平距離'] * 10000)
                today_count += 1
                second_x = x_count
                flat_sum += (item.get('水平距離', 0) or 0) + (item.get('水平距離2', 0) or 0)
                haikann_sum += (item.get('配管長', 0) or 0) + (item.get('配管長2', 0) or 0)
                msp.add_line((first_x, y_count - y_count - 45000000), (first_x, y_count - y_count - 61000000))
                msp.add_line((second_x, y_count - y_count - 45000000), (second_x, y_count - y_count - 61000000))
                msp.add_line((first_x, y_count - y_count - 61000000), (second_x, y_count - y_count - 61000000))
                cn.flat_line1(msp, item, offset_x=x_count - (item['水平距離'] * 10000 / 2), offset_y=y_count - y_count - 48000000)
                if item['水平距離'] == item['配管長']:
                    pass
                else:
                    cn.haikann_line1(msp, item, offset_x=x_count -9200000 - (item['水平距離'] * 10000 / 2), offset_y=y_count - y_count - 42000000)        
    elif (
        '直管' in name
        and '（乙切管）' in name
        and '(V)' in name
        and down_22
    ):
        if item['方向'] == 'BB':
            if today_count == 0:
                cn.vb_tyokkann_otu_22_11_down_bb(msp, item, offset_x=x_count, offset_y=y_count)
                x_count += (item['水平距離'] * 10000)
                y_count += - ((item['配管長'] * 10000) ** 2 - (item['水平距離'] * 10000) ** 2) ** 0.5
                msp.add_text(f"乙{item['配管長']}", dxfattribs={
                'insert': (x_count - 6000000 - (item['水平距離'] * 10000 / 2),
                            y_count - 11000000 - ((item['配管長'] * 10000) ** 2
                                                  - (item['水平距離'] * 10000) ** 2) ** 0.5),
                'height': 6000000,
                })
                today_xbase += (item['水平距離'] * 10000)
                today_count += 1
            else:
                cn.choose_zairyou_no(msp, item, offset_x=x_count + (item['水平距離'] * 10000 / 2), offset_y=y_count + 5000000 + (item['配管長'] * 10000))                
                cn.vb_tyokkann_otu_22_11_down_bb(msp, item, offset_x=x_count, offset_y=y_count)
                msp.add_text(f"乙{item['配管長']}", dxfattribs={
                'insert': (x_count - 6000000 - (item['水平距離'] * 10000 / 2),
                            y_count - 11000000 - ((item['配管長'] * 10000) ** 2
                                                  - (item['水平距離'] * 10000) ** 2) ** 0.5),
                'height': 6000000,
                })
                first_x = x_count
                x_count += (item['水平距離'] * 10000)
                y_count += - ((item['配管長'] * 10000) ** 2 - (item['水平距離'] * 10000) ** 2) ** 0.5
                today_count += 1
                second_x = x_count
                flat_sum += (item.get('水平距離', 0) or 0) + (item.get('水平距離2', 0) or 0)
                haikann_sum += (item.get('配管長', 0) or 0) + (item.get('配管長2', 0) or 0)
                msp.add_line((first_x, y_count - y_count - 45000000), (first_x, y_count - y_count - 61000000))
                msp.add_line((second_x, y_count - y_count - 45000000), (second_x, y_count - y_count - 61000000))
                msp.add_line((first_x, y_count - y_count - 61000000), (second_x, y_count - y_count - 61000000))
                cn.flat_line1(msp, item, offset_x=x_count - (item['水平距離'] * 10000 / 2), offset_y=y_count - y_count - 48000000)
                if item['水平距離'] == item['配管長']:
                    pass
                else:
                    cn.haikann_line1(msp, item, offset_x=x_count -9200000 - (item['水平距離'] * 10000 / 2), offset_y=y_count - y_count - 42000000)

    elif (
        '直管' in name
        and '（乙切管）' in name
        and '(V)' in name
        and up_22
    ):
        if item['方向'] == 'BB':
            if today_count == 0:
                cn.vb_tyokkann_otu_22_11_up_bb(msp, item, offset_x=x_count, offset_y=y_count)
                x_count += (item['水平距離'] * 10000)
                y_count += ((item['配管長'] * 10000) ** 2 - (item['水平距離'] * 10000) ** 2) ** 0.5
                msp.add_text(f"乙{item['配管長']}", dxfattribs={
                'insert': (x_count - (item['水平距離'] * 10000 / 2),
                            y_count - 11000000 - ((item['配管長'] * 10000) ** 2
                                                  - (item['水平距離'] * 10000) ** 2) ** 0.5),
                'height': 6000000,
                })
                today_xbase += (item['水平距離'] * 10000)
                today_count += 1
            else:
                cn.choose_zairyou_no(msp, item, offset_x=x_count + (item['水平距離'] * 10000 / 2), offset_y=y_count + 5000000 + (item['配管長'] * 10000))                
                cn.vb_tyokkann_otu_22_11_up_bb(msp, item, offset_x=x_count, offset_y=y_count)
                msp.add_text(f"乙{item['配管長']}", dxfattribs={
                'insert': (x_count - (item['水平距離'] * 10000 / 2),
                            y_count - 11000000 - ((item['配管長'] * 10000) ** 2
                                                  - (item['水平距離'] * 10000) ** 2) ** 0.5),
                'height': 6000000,
                })
                first_x = x_count
                x_count += (item['水平距離'] * 10000)
                y_count += ((item['配管長'] * 10000) ** 2 - (item['水平距離'] * 10000) ** 2) ** 0.5
                today_count += 1
                second_x = x_count
                flat_sum += (item.get('水平距離', 0) or 0) + (item.get('水平距離2', 0) or 0)
                haikann_sum += (item.get('配管長', 0) or 0) + (item.get('配管長2', 0) or 0)
                msp.add_line((first_x, y_count - y_count - 45000000), (first_x, y_count - y_count - 61000000))
                msp.add_line((second_x, y_count - y_count - 45000000), (second_x, y_count - y_count - 61000000))
                msp.add_line((first_x, y_count - y_count - 61000000), (second_x, y_count - y_count - 61000000))
                cn.flat_line1(msp, item, offset_x=x_count - (item['水平距離'] * 10000 / 2), offset_y=y_count - y_count - 48000000)
                if item['水平距離'] == item['配管長']:
                    pass
                else:
                    cn.haikann_line1(msp, item, offset_x=x_count -9200000 - (item['水平距離'] * 10000 / 2), offset_y=y_count - y_count - 42000000)

    elif (
        '直管' in name
        and '（乙切管）' in name
        and '(V)' in name
        and down_11
    ):
        if item['方向'] == 'BB':
            if today_count == 0:
                cn.vb_tyokkann_otu_22_11_down_bb(msp, item, offset_x=x_count, offset_y=y_count)
                x_count += (item['水平距離'] * 10000)
                y_count += - ((item['配管長'] * 10000) ** 2 - (item['水平距離'] * 10000) ** 2) ** 0.5
                msp.add_text(f"乙{item['配管長']}", dxfattribs={
                'insert': (x_count - 6000000 - (item['水平距離'] * 10000 / 2),
                            y_count - 11000000 - ((item['配管長'] * 10000) ** 2
                                                  - (item['水平距離'] * 10000) ** 2) ** 0.5),
                'height': 6000000,
                })
                today_xbase += (item['水平距離'] * 10000)
                today_count += 1
            else:
                cn.choose_zairyou_no(msp, item, offset_x=x_count + (item['水平距離'] * 10000 / 2), offset_y=y_count + 5000000 + (item['配管長'] * 10000))                
                cn.vb_tyokkann_otu_22_11_down_bb(msp, item, offset_x=x_count, offset_y=y_count)
                msp.add_text(f"乙{item['配管長']}", dxfattribs={
                'insert': (x_count - 6000000 - (item['水平距離'] * 10000 / 2),
                            y_count - 11000000 - ((item['配管長'] * 10000) ** 2
                                                  - (item['水平距離'] * 10000) ** 2) ** 0.5),
                'height': 6000000,
                })
                first_x = x_count
                x_count += (item['水平距離'] * 10000)
                y_count += - ((item['配管長'] * 10000) ** 2 - (item['水平距離'] * 10000) ** 2) ** 0.5
                today_count += 1
                second_x = x_count
                flat_sum += (item.get('水平距離', 0) or 0) + (item.get('水平距離2', 0) or 0)
                haikann_sum += (item.get('配管長', 0) or 0) + (item.get('配管長2', 0) or 0)
                msp.add_line((first_x, y_count - y_count - 45000000), (first_x, y_count - y_count - 61000000))
                msp.add_line((second_x, y_count - y_count - 45000000), (second_x, y_count - y_count - 61000000))
                msp.add_line((first_x, y_count - y_count - 61000000), (second_x, y_count - y_count - 61000000))
                cn.flat_line1(msp, item, offset_x=x_count - (item['水平距離'] * 10000 / 2), offset_y=y_count - y_count - 48000000)
                if item['水平距離'] == item['配管長']:
                    pass
                else:
                    cn.haikann_line1(msp, item, offset_x=x_count -9200000 - (item['水平距離'] * 10000 / 2), offset_y=y_count - y_count - 42000000)

    elif (
        '直管' in name
        and '（乙切管）' in name
        and '(V)' in name
        and up_11
    ):
        if item['方向'] == 'BB':
            if today_count == 0:
                cn.vb_tyokkann_otu_22_11_up_bb(msp, item, offset_x=x_count, offset_y=y_count)
                x_count += (item['水平距離'] * 10000)
                y_count += ((item['配管長'] * 10000) ** 2 - (item['水平距離'] * 10000) ** 2) ** 0.5
                msp.add_text(f"乙{item['配管長']}", dxfattribs={
                'insert': (x_count - (item['水平距離'] * 10000 / 2),
                            y_count - 11000000 - ((item['配管長'] * 10000) ** 2
                                                  - (item['水平距離'] * 10000) ** 2) ** 0.5),
                'height': 6000000,
                })
                today_xbase += (item['水平距離'] * 10000)
                today_count += 1
            else:
                cn.choose_zairyou_no(msp, item, offset_x=x_count + (item['水平距離'] * 10000 / 2), offset_y=y_count + 5000000 + (item['配管長'] * 10000))                
                cn.vb_tyokkann_otu_22_11_up_bb(msp, item, offset_x=x_count, offset_y=y_count)
                msp.add_text(f"乙{item['配管長']}", dxfattribs={
                'insert': (x_count - (item['水平距離'] * 10000 / 2),
                            y_count - 11000000 - ((item['配管長'] * 10000) ** 2
                                                  - (item['水平距離'] * 10000) ** 2) ** 0.5),
                'height': 6000000,
                })
                first_x = x_count
                x_count += (item['水平距離'] * 10000)
                y_count += ((item['配管長'] * 10000) ** 2 - (item['水平距離'] * 10000) ** 2) ** 0.5
                today_count += 1
                second_x = x_count
                flat_sum += (item.get('水平距離', 0) or 0) + (item.get('水平距離2', 0) or 0)
                haikann_sum += (item.get('配管長', 0) or 0) + (item.get('配管長2', 0) or 0)
                msp.add_line((first_x, y_count - y_count - 45000000), (first_x, y_count - y_count - 61000000))
                msp.add_line((second_x, y_count - y_count - 45000000), (second_x, y_count - y_count - 61000000))
                msp.add_line((first_x, y_count - y_count - 61000000), (second_x, y_count - y_count - 61000000))
                cn.flat_line1(msp, item, offset_x=x_count - (item['水平距離'] * 10000 / 2), offset_y=y_count - y_count - 48000000)
                if item['水平距離'] == item['配管長']:
                    pass
                else:
                    cn.haikann_line1(msp, item, offset_x=x_count -9200000 - (item['水平距離'] * 10000 / 2), offset_y=y_count - y_count - 42000000)

    elif (
        '直管' in name
        and '（乙切管）' in name
    ):
        if item['方向'] == 'BB':
            if today_count == 0:
                cn.tyokkann_otsu_bb(msp, item, offset_x=x_count, offset_y=y_count)
                x_count += (item['配管長'] * 10000) / 2
                msp.add_text(f"乙{item['配管長']}", dxfattribs={
                'insert': (x_count - 7770000 + (item['配管長'] * 10000) / 2, y_count - 12000000),
                'height': 6000000,
                })
                today_xbase += (item['配管長'] * 10000) / 2
                today_count += 1
            else:
                cn.choose_zairyou_no(msp, item, offset_x=x_count + (item['配管長'] * 10000 / 4), offset_y=y_count + 15000000)
                cn.tyokkann_otsu_bb(msp, item, offset_x=x_count, offset_y=y_count)
                msp.add_text(f"乙{item['配管長']}", dxfattribs={
                'insert': (x_count - 7770000 + (item['配管長'] * 10000) / 4, y_count - 12000000),
                'height': 6000000,
                })
                first_x = x_count
                x_count += (item['配管長'] * 10000) / 2
                today_count += 1
                second_x = x_count
                flat_sum += (item.get('水平距離', 0) or 0) + (item.get('水平距離2', 0) or 0)
                haikann_sum += (item.get('配管長', 0) or 0) + (item.get('配管長2', 0) or 0)
                msp.add_line((first_x, y_count - y_count - 45000000), (first_x, y_count - y_count - 61000000))
                msp.add_line((second_x, y_count - y_count - 45000000), (second_x, y_count - y_count - 61000000))
                msp.add_line((first_x, y_count - y_count - 61000000), (second_x, y_count - y_count - 61000000))
                cn.flat_line1(msp, item, offset_x=x_count + 6850000 - (item['配管長'] * 10000 / 2), offset_y=y_count - y_count - 61000000) 
    elif (
        'GX形ダクタイル鋳鉄' in name
    ):
        if item['方向'] == 'BA':
            cn.gx_tyokkann_ba(msp, offset_x=x_count, offset_y=y_count)
            x_count += 39750000
        elif item['方向'] == 'AB':
            cn.gx_tyokkann_ab(msp, offset_x=x_count, offset_y=y_count)
            x_count += 39750000

    elif (
        'GX形' in name
        and '45°' in name
        and '(V)' in name
        and '下がり' in name
    ):
        if item['方向'] == 'AA':
            if today_count == 0:
                cn.gx_vb_ryouuke_bend_45_down_st_aa(msp, offset_x=x_count, offset_y=y_count)
                x_count += 33250000
                y_count += -13750000
                today_count += 1
                today_xbase += 33250000
            else:
                cn.choose_zairyou_no(msp, item, offset_x=x_count + 16625000, offset_y=y_count + 15000000)
                cn.name_sokutenn(msp, item, route_name,offset_x=x_count - 90000000, offset_y=y_count - y_count + 13500000 * y_offsetcount)
                cn.gx_vb_ryouuke_bend_45_down_st_aa(msp, offset_x=x_count, offset_y=y_count)
                first_x = x_count
                x_count += 33250000
                y_count += -13750000
                y_offsetcount += 1
                today_count += 1
                second_x = x_count
                flat_sum += (item.get('水平距離', 0) or 0) + (item.get('水平距離2', 0) or 0)
                haikann_sum += (item.get('配管長', 0) or 0) + (item.get('配管長2', 0) or 0)
                msp.add_line((first_x, y_count - y_count - 45000000), (first_x, y_count - y_count - 61000000))
                msp.add_line((first_x + 19525000, y_count - y_count - 45000000), (first_x + 19525000, y_count - y_count - 61000000))
                msp.add_line((second_x, y_count - y_count - 45000000), (second_x, y_count - y_count - 61000000))
                msp.add_line((first_x, y_count - y_count - 61000000), (second_x, y_count - y_count - 61000000))
                cn.flat_line1(msp, item, offset_x=x_count-23470000, offset_y=y_count - y_count - 61000000)
                cn.flat_line2(msp, item, offset_x=x_count-5405000, offset_y=y_count - y_count - 61000000)
                if item['水平距離'] == item['配管長']:
                    pass
                else:
                    cn.haikann_line1(msp, item, offset_x=x_count-19500000, offset_y=y_count - y_count - 54500000)                   
                if item['水平距離2'] == item['配管長2']:
                    pass
                else:
                    cn.haikann_line2(msp, item, offset_x=x_count-14350000, offset_y=y_count - y_count - 54500000)
        elif item['方向'] == 'AB':
            if today_count == 0:
                cn.gx_vb_katauke_bend_45_down_st_ab(msp, offset_x=x_count, offset_y=y_count)
                x_count += 25160000
                y_count += -5660000
                today_count += 1
                today_xbase += 25160000
            else:
                cn.choose_zairyou_no(msp, item, offset_x=x_count + 12580000, offset_y=y_count + 15000000)
                cn.name_sokutenn(msp, item, route_name,offset_x=x_count - 90000000, offset_y=y_count - y_count + 13500000 * y_offsetcount)
                cn.gx_vb_katauke_bend_45_down_st_ab(msp, offset_x=x_count, offset_y=y_count)
                first_x = x_count
                x_count += 25160000
                y_count += -5660000
                y_offsetcount += 1
                today_count += 1
                second_x = x_count
                flat_sum += (item.get('水平距離', 0) or 0) + (item.get('水平距離2', 0) or 0)
                haikann_sum += (item.get('配管長', 0) or 0) + (item.get('配管長2', 0) or 0)
                msp.add_line((first_x, y_count - y_count - 45000000), (first_x, y_count - y_count - 61000000))
                msp.add_line((first_x + 10000000, y_count - y_count - 45000000), (first_x + 10000000, y_count - y_count - 61000000))
                msp.add_line((second_x, y_count - y_count - 45000000), (second_x, y_count - y_count - 61000000))
                msp.add_line((first_x, y_count - y_count - 61000000), (second_x, y_count - y_count - 61000000))
                cn.flat_line1(msp, item, offset_x=x_count-20200000, offset_y=y_count - y_count - 61000000)
                cn.flat_line2(msp, item, offset_x=x_count-7650000, offset_y=y_count - y_count - 61000000)
                if item['水平距離'] == item['配管長']:
                    pass
                else:
                    cn.haikann_line1(msp, item, offset_x=x_count-19500000, offset_y=y_count - y_count - 54500000)                   
                if item['水平距離2'] == item['配管長2']:
                    pass
                else:
                    cn.haikann_line2(msp, item, offset_x=x_count-15150000, offset_y=y_count - y_count - 54500000)
        elif item['方向'] == 'BA':
            if today_count == 0:
                cn.gx_vb_katauke_bend_45_down_st_ba(msp, offset_x=x_count, offset_y=y_count)
                x_count += 21790000
                y_count += - 13750000
                today_count += 1
                today_xbase += 21790000
            else:
                cn.choose_zairyou_no(msp, item, offset_x=x_count + 8660000, offset_y=y_count + 15000000)
                cn.name_sokutenn(msp, item, route_name,offset_x=x_count - 90000000, offset_y=y_count - y_count + 13500000 * y_offsetcount)
                cn.gx_vb_katauke_bend_45_down_st_ba(msp, offset_x=x_count, offset_y=y_count)
                first_x = x_count
                x_count += 21790000
                y_count += -13750000
                y_offsetcount += 1
                today_count += 1
                second_x = x_count
                flat_sum += (item.get('水平距離', 0) or 0) + (item.get('水平距離2', 0) or 0)
                haikann_sum += (item.get('配管長', 0) or 0) + (item.get('配管長2', 0) or 0)
                msp.add_line((first_x, y_count - y_count - 45000000), (first_x, y_count - y_count - 61000000))
                msp.add_line((first_x + 14500000, y_count - y_count - 45000000), (first_x + 14500000, y_count - y_count - 61000000))
                msp.add_line((second_x, y_count - y_count - 45000000), (second_x, y_count - y_count - 61000000))
                msp.add_line((first_x, y_count - y_count - 61000000), (second_x, y_count - y_count - 61000000))
                cn.flat_line1(msp, item, offset_x=x_count-14600000, offset_y=y_count - y_count - 61000000)
                cn.flat_line2(msp, item, offset_x=x_count-2150000, offset_y=y_count - y_count - 61000000)
                if item['水平距離'] == item['配管長']:
                    pass
                else:
                    cn.haikann_line1(msp, item, offset_x=x_count-19500000, offset_y=y_count - y_count - 54500000)                   
                if item['水平距離2'] == item['配管長2']:
                    pass
                else:
                    cn.haikann_line2(msp, item, offset_x=x_count-11000000, offset_y=y_count - y_count - 54500000)    
    elif (
        'GX形' in name
        and '45°' in name
        and '(V)' in name
        and '上がり' in name
    ):
        if item['方向'] == 'AA':
            if today_count == 0:
                cn.gx_vb_ryouuke_bend_45_up_st_aa(msp, offset_x=x_count, offset_y=y_count)
                x_count += 33250000
                y_count += 13750000
                today_count += 1
                today_xbase += 33250000
            else:
                cn.choose_zairyou_no(msp, item, offset_x=x_count + 16850000, offset_y=y_count + 15000000)
                cn.name_sokutenn(msp, item, route_name,offset_x=x_count - 90000000, offset_y=y_count - y_count + 13500000 * y_offsetcount)
                cn.gx_vb_ryouuke_bend_45_up_st_aa(msp, offset_x=x_count, offset_y=y_count)
                first_x = x_count
                x_count += 33250000
                y_count += 13750000
                y_offsetcount += 1
                today_count += 1
                second_x = x_count
                flat_sum += (item.get('水平距離', 0) or 0) + (item.get('水平距離2', 0) or 0)
                haikann_sum += (item.get('配管長', 0) or 0) + (item.get('配管長2', 0) or 0)
                msp.add_line((first_x, y_count - y_count - 45000000), (first_x, y_count - y_count - 61000000))
                msp.add_line((first_x + 19500000, y_count - y_count - 45000000), (first_x + 19500000, y_count - y_count - 61000000))
                msp.add_line((second_x, y_count - y_count - 45000000), (second_x, y_count - y_count - 61000000))
                msp.add_line((first_x, y_count - y_count - 61000000), (second_x, y_count - y_count - 61000000))
                cn.flat_line1(msp, item, offset_x=x_count-23600000, offset_y=y_count - y_count - 61000000)
                cn.flat_line2(msp, item, offset_x=x_count-4700000, offset_y=y_count - y_count - 61000000)
                if item['水平距離'] == item['配管長']:
                    pass
                else:
                    cn.haikann_line1(msp, item, offset_x=x_count-19500000, offset_y=y_count - y_count - 54500000)                   
                if item['水平距離2'] == item['配管長2']:
                    pass
                else:
                    cn.haikann_line2(msp, item, offset_x=x_count-14450000, offset_y=y_count - y_count - 54500000)
        elif item['方向'] == 'AB':
            if today_count == 0:
                cn.gx_vb_katauke_bend_45_up_st_ab(msp, offset_x=x_count, offset_y=y_count)
                x_count += 25160000
                y_count += 5660000
                today_count += 1
                today_xbase += 25160000
            else:
                cn.choose_zairyou_no(msp, item, offset_x=x_count + 17000000, offset_y=y_count + 15000000)
                cn.name_sokutenn(msp, item, route_name,offset_x=x_count - 90000000, offset_y=y_count - y_count + 13500000 * y_offsetcount)
                cn.gx_vb_katauke_bend_45_up_st_ab(msp, offset_x=x_count, offset_y=y_count)
                first_x = x_count
                x_count += 25160000
                y_count += 5660000
                y_offsetcount += 1
                today_count += 1
                second_x = x_count
                flat_sum += (item.get('水平距離', 0) or 0) + (item.get('水平距離2', 0) or 0)
                haikann_sum += (item.get('配管長', 0) or 0) + (item.get('配管長2', 0) or 0)
                msp.add_line((first_x, y_count - y_count - 45000000), (first_x, y_count - y_count - 61000000))
                msp.add_line((first_x + 10000000, y_count - y_count - 45000000), (first_x + 10000000, y_count - y_count - 61000000))
                msp.add_line((second_x, y_count - y_count - 45000000), (second_x, y_count - y_count - 61000000))
                msp.add_line((first_x, y_count - y_count - 61000000), (second_x, y_count - y_count - 61000000))
                cn.flat_line1(msp, item, offset_x=x_count-20250000, offset_y=y_count - y_count - 61000000)
                cn.flat_line2(msp, item, offset_x=x_count-6100000, offset_y=y_count - y_count - 61000000)
                if item['水平距離'] == item['配管長']:
                    pass
                else:
                    cn.haikann_line1(msp, item, offset_x=x_count-19500000, offset_y=y_count - y_count - 54500000)                   
                if item['水平距離2'] == item['配管長2']:
                    pass
                else:
                    cn.haikann_line2(msp, item, offset_x=x_count-15150000, offset_y=y_count - y_count - 54500000)
        elif item['方向'] == 'BA':
            if today_count == 0:
                cn.gx_vb_katauke_bend_45_up_st_ba(msp, offset_x=x_count, offset_y=y_count)
                x_count += 21790000
                y_count += 13750000
                today_count += 1
                today_xbase += 21790000
            else:
                cn.choose_zairyou_no(msp, item, offset_x=x_count + 5000000, offset_y=y_count + 15000000)
                cn.name_sokutenn(msp, item, route_name,offset_x=x_count - 90000000, offset_y=y_count - y_count + 13500000 * y_offsetcount)
                cn.gx_vb_katauke_bend_45_up_st_ba(msp, offset_x=x_count, offset_y=y_count)
                first_x = x_count
                x_count += 21790000
                y_count += 13750000
                y_offsetcount += 1
                today_count += 1
                second_x = x_count
                flat_sum += (item.get('水平距離', 0) or 0) + (item.get('水平距離2', 0) or 0)
                haikann_sum += (item.get('配管長', 0) or 0) + (item.get('配管長2', 0) or 0)
                msp.add_line((first_x, y_count - y_count - 45000000), (first_x, y_count - y_count - 61000000))
                msp.add_line((first_x + 15000000, y_count - y_count - 45000000), (first_x + 15000000, y_count - y_count - 61000000))
                msp.add_line((second_x, y_count - y_count - 45000000), (second_x, y_count - y_count - 61000000))
                msp.add_line((first_x, y_count - y_count - 61000000), (second_x, y_count - y_count - 61000000))
                cn.flat_line1(msp, item, offset_x=x_count-14350000, offset_y=y_count - y_count - 61000000)
                cn.flat_line2(msp, item, offset_x=x_count-1900000, offset_y=y_count - y_count - 61000000)
                if item['水平距離'] == item['配管長']:
                    pass
                else:
                    cn.haikann_line1(msp, item, offset_x=x_count-19500000, offset_y=y_count - y_count - 54500000)                   
                if item['水平距離2'] == item['配管長2']:
                    pass
                else:
                    cn.haikann_line2(msp, item, offset_x=x_count-10850000, offset_y=y_count - y_count - 54500000)
    elif (
        'GX形ダクタイル鋳鉄' in name
        and '曲管' in name
        and '22.5°' in name
        and '(V)' in name
        and '下がり' in name
    ):
        if item['方向'] == 'AA':
            cn.gx_vb_ryouuke_bend_22_down_st_aa(msp, offset_x=x_count, offset_y=y_count)
            x_count += 37510000
            y_count += - 7460000
    elif (
        'GX形ダクタイル鋳鉄' in name
        and '曲管' in name
        and '22.5°' in name
        and '(V)' in name
        and '上がり' in name
    ):
        if item['方向'] == 'AA':
            cn.gx_vb_ryouuke_bend_22_up_st_aa(msp, offset_x=x_count, offset_y=y_count)
            x_count += 37510000
            y_count += 7460000
    elif (
        'GX形ダクタイル鋳鉄' in name
        and '曲管' in name
        and '11.25°' in name
        and '(V)' in name
        and '下がり' in name
    ):
        if item['方向'] == 'AA':
            cn.gx_vb_ryouuke_bend_11_down_st_aa(msp, offset_x=x_count, offset_y=y_count)
            x_count += 38630000
            y_count += - 3810000
    elif (
        'GX形ダクタイル鋳鉄' in name
        and '曲管' in name
        and '11.25°' in name
        and '(V)' in name
        and '上がり' in name
    ):
        if item['方向'] == 'AA':
            cn.gx_vb_ryouuke_bend_11_up_st_aa(msp, offset_x=x_count, offset_y=y_count)
            x_count += 38630000
            y_count += 3810000
    elif (
        'GX形' in name
        and '45°' in name
        and '（V）' in name
        and down_45
    ):
        if item['方向'] == 'AA':
            if today_count == 0:
                cn.gx_vb_ryouuke_bend_45_down_ed_aa(msp, offset_x=x_count, offset_y=y_count)
                x_count += 33290000
                y_count += -13790000
                today_count += 1
                today_xbase += 33290000
            else:
                cn.choose_zairyou_no(msp, item, offset_x=x_count + 18500000, offset_y=y_count + 3000000)
                cn.name_sokutenn(msp, item, route_name, offset_x=x_count - 90000000, offset_y=y_count - y_count + 13500000 * y_offsetcount)
                cn.gx_vb_ryouuke_bend_45_down_ed_aa(msp, offset_x=x_count, offset_y=y_count)
                first_x = x_count
                x_count += 33290000
                y_count += -13790000
                y_offsetcount += 1
                today_count += 1
                second_x = x_count
                flat_sum += (item.get('水平距離', 0) or 0) + (item.get('水平距離2', 0) or 0)
                haikann_sum += (item.get('配管長', 0) or 0) + (item.get('配管長2', 0) or 0)
                msp.add_line((first_x, y_count - y_count - 45000000), (first_x, y_count - y_count - 61000000))
                msp.add_line((second_x, y_count - y_count - 45000000), (second_x, y_count - y_count - 61000000))
                msp.add_line((second_x - 19500000, y_count - y_count - 45000000), (second_x - 19500000, y_count - y_count - 61000000))
                msp.add_line((first_x, y_count - y_count - 61000000), (second_x, y_count - y_count - 61000000))
                cn.flat_line1(msp, item, offset_x=x_count-24850000, offset_y=y_count - y_count - 61000000)
                cn.flat_line2(msp, item, offset_x=x_count-9750000, offset_y=y_count - y_count - 61000000)
                if item['水平距離'] == item['配管長']:
                    pass
                else:
                    cn.haikann_line1(msp, item, offset_x=x_count-33900000, offset_y=y_count - y_count - 54500000)                
                if item['水平距離2'] == item['配管長2']:
                    pass
                else:
                    cn.haikann_line2(msp, item, offset_x=x_count-12600000, offset_y=y_count - y_count - 54500000)
        elif item['方向'] == 'AB':
            if today_count == 0:
                cn.gx_vb_katauke_bend_45_down_ed_ab(msp, offset_x=x_count, offset_y=y_count)
                x_count += 21790000
                y_count += -13790000
                today_count += 1
                today_xbase += 21790000
            else:
                cn.choose_zairyou_no(msp, item, offset_x=x_count + 17800000, offset_y=y_count + 3000000)
                cn.name_sokutenn(msp, item, route_name, offset_x=x_count - 90000000, offset_y=y_count - y_count + 13500000 * y_offsetcount)
                cn.gx_vb_katauke_bend_45_down_ed_ab(msp, offset_x=x_count, offset_y=y_count)
                first_x = x_count
                x_count += 21790000
                y_count += -13790000
                y_offsetcount += 1
                today_count += 1
                second_x = x_count
                flat_sum += (item.get('水平距離', 0) or 0) + (item.get('水平距離2', 0) or 0)
                haikann_sum += (item.get('配管長', 0) or 0) + (item.get('配管長2', 0) or 0)
                msp.add_line((first_x, y_count - y_count - 45000000), (first_x, y_count - y_count - 61000000))
                msp.add_line((second_x, y_count - y_count - 45000000), (second_x, y_count - y_count - 61000000))
                msp.add_line((second_x - 14500000, y_count - y_count - 45000000), (second_x - 14500000, y_count - y_count - 61000000))
                msp.add_line((first_x, y_count - y_count - 61000000), (second_x, y_count - y_count - 61000000))
                cn.flat_line1(msp, item, offset_x=x_count-16600000, offset_y=y_count - y_count - 61000000)
                cn.flat_line2(msp, item, offset_x=x_count-7300000, offset_y=y_count - y_count - 61000000)
                if item['水平距離'] == item['配管長']:
                    pass
                else:
                    cn.haikann_line1(msp, item, offset_x=x_count-25650000, offset_y=y_count - y_count - 54500000)                
                if item['水平距離2'] == item['配管長2']:
                    pass
                else:
                    cn.haikann_line2(msp, item, offset_x=x_count-12600000, offset_y=y_count - y_count - 54500000)
        elif item['方向'] == 'BA':
            if today_count == 0:
                cn.gx_vb_katauke_bend_45_down_ed_ba(msp, offset_x=x_count, offset_y=y_count)
                x_count += 25160000
                y_count += -5660000
                today_count += 1
                today_xbase += 25160000
            else:
                cn.choose_zairyou_no(msp, item, offset_x=x_count + 14200000, offset_y=y_count + 8000000)
                cn.name_sokutenn(msp, item, route_name, offset_x=x_count - 90000000, offset_y=y_count - y_count + 13500000 * y_offsetcount)
                cn.gx_vb_katauke_bend_45_down_ed_ba(msp, offset_x=x_count, offset_y=y_count)
                first_x = x_count
                x_count += 25160000
                y_count += -5660000
                y_offsetcount += 1
                today_count += 1
                second_x = x_count
                flat_sum += (item.get('水平距離', 0) or 0) + (item.get('水平距離2', 0) or 0)
                haikann_sum += (item.get('配管長', 0) or 0) + (item.get('配管長2', 0) or 0)
                msp.add_line((first_x, y_count - y_count - 45000000), (first_x, y_count - y_count - 61000000))
                msp.add_line((second_x, y_count - y_count - 45000000), (second_x, y_count - y_count - 61000000))
                msp.add_line((second_x - 10000000, y_count - y_count - 45000000), (second_x - 10000000, y_count - y_count - 61000000))
                msp.add_line((first_x, y_count - y_count - 61000000), (second_x, y_count - y_count - 61000000))
                cn.flat_line1(msp, item, offset_x=x_count-16050000, offset_y=y_count - y_count - 61000000)
                cn.flat_line2(msp, item, offset_x=x_count-5000000, offset_y=y_count - y_count - 61000000)
                if item['水平距離'] == item['配管長']:
                    pass
                else:
                    cn.haikann_line1(msp, item, offset_x=x_count-25100000, offset_y=y_count - y_count - 54500000)                
                if item['水平距離2'] == item['配管長2']:
                    pass
                else:
                    cn.haikann_line2(msp, item, offset_x=x_count-12600000, offset_y=y_count - y_count - 54500000)
    elif (
        'GX形' in name
        and '45°' in name
        and '（V）' in name
        and up_45
    ):
        if item['方向'] == 'AA':
            if today_count == 0:
                cn.gx_vb_ryouuke_bend_45_up_ed_aa(msp, offset_x=x_count, offset_y=y_count)
                x_count += 33290000
                y_count += 13790000
                today_count += 1
                today_xbase += 33290000
            else:
                cn.choose_zairyou_no(msp, item, offset_x=x_count + 16300000, offset_y=y_count + 22000000)
                cn.name_sokutenn(msp, item, route_name, offset_x=x_count - 90000000, offset_y=y_count - y_count + 13500000 * y_offsetcount)
                cn.gx_vb_ryouuke_bend_45_up_ed_aa(msp, offset_x=x_count, offset_y=y_count)
                first_x = x_count
                x_count += 33290000
                y_count += 13790000
                y_offsetcount += 1
                today_count += 1
                second_x = x_count
                flat_sum += (item.get('水平距離', 0) or 0) + (item.get('水平距離2', 0) or 0)
                haikann_sum += (item.get('配管長', 0) or 0) + (item.get('配管長2', 0) or 0)
                msp.add_line((first_x, y_count - y_count - 45000000), (first_x, y_count - y_count - 61000000))
                msp.add_line((second_x, y_count - y_count - 45000000), (second_x, y_count - y_count - 61000000))
                msp.add_line((second_x - 19500000, y_count - y_count - 45000000), (second_x - 19500000, y_count - y_count - 61000000))
                msp.add_line((first_x, y_count - y_count - 61000000), (second_x, y_count - y_count - 61000000))
                cn.flat_line1(msp, item, offset_x=x_count-24850000, offset_y=y_count - y_count - 61000000)
                cn.flat_line2(msp, item, offset_x=x_count-9750000, offset_y=y_count - y_count - 61000000)
                if item['水平距離'] == item['配管長']:
                    pass
                else:
                    cn.haikann_line1(msp, item, offset_x=x_count-33900000, offset_y=y_count - y_count - 54500000)                
                if item['水平距離2'] == item['配管長2']:
                    pass
                else:
                    cn.haikann_line2(msp, item, offset_x=x_count-12600000, offset_y=y_count - y_count - 54500000)
        elif item['方向'] == 'AB':
            if today_count == 0:
                cn.gx_vb_katauke_bend_45_up_ed_ab(msp, offset_x=x_count, offset_y=y_count)
                x_count += 21790000
                y_count += 13790000
                today_count += 1
                today_xbase += 21790000
            else:
                cn.choose_zairyou_no(msp, item, offset_x=x_count + 8500000, offset_y=y_count + 22500000)
                cn.name_sokutenn(msp, item, route_name, offset_x=x_count - 90000000, offset_y=y_count - y_count + 13500000 * y_offsetcount)
                cn.gx_vb_katauke_bend_45_up_ed_ab(msp, offset_x=x_count, offset_y=y_count)
                first_x = x_count
                x_count += 21790000
                y_count += 13790000
                y_offsetcount += 1
                today_count += 1
                second_x = x_count
                flat_sum += (item.get('水平距離', 0) or 0) + (item.get('水平距離2', 0) or 0)
                haikann_sum += (item.get('配管長', 0) or 0) + (item.get('配管長2', 0) or 0)
                msp.add_line((first_x, y_count - y_count - 45000000), (first_x, y_count - y_count - 61000000))
                msp.add_line((second_x, y_count - y_count - 45000000), (second_x, y_count - y_count - 61000000))
                msp.add_line((second_x - 15000000, y_count - y_count - 45000000), (second_x - 15000000, y_count - y_count - 61000000))
                msp.add_line((first_x, y_count - y_count - 61000000), (second_x, y_count - y_count - 61000000))
                cn.flat_line1(msp, item, offset_x=x_count-16800000, offset_y=y_count - y_count - 61000000)
                cn.flat_line2(msp, item, offset_x=x_count-7500000, offset_y=y_count - y_count - 61000000)
                if item['水平距離'] == item['配管長']:
                    pass
                else:
                    cn.haikann_line1(msp, item, offset_x=x_count-25950000, offset_y=y_count - y_count - 54500000)                
                if item['水平距離2'] == item['配管長2']:
                    pass
                else:
                    cn.haikann_line2(msp, item, offset_x=x_count-12600000, offset_y=y_count - y_count - 54500000)
        elif item['方向'] == 'BA':
            if today_count == 0:
                cn.gx_vb_katauke_bend_45_up_ed_ba(msp, offset_x=x_count, offset_y=y_count)
                x_count += 25160000
                y_count += 5660000
                today_count += 1
                today_xbase += 25160000
            else:
                cn.choose_zairyou_no(msp, item, offset_x=x_count + 8500000, offset_y=y_count + 15000000)
                cn.name_sokutenn(msp, item, route_name, offset_x=x_count - 90000000, offset_y=y_count - y_count + 13500000 * y_offsetcount)
                cn.gx_vb_katauke_bend_45_up_ed_ba(msp, offset_x=x_count, offset_y=y_count)
                first_x = x_count
                x_count += 25160000
                y_count += 5660000
                y_offsetcount += 1
                today_count += 1
                second_x = x_count
                flat_sum += (item.get('水平距離', 0) or 0) + (item.get('水平距離2', 0) or 0)
                haikann_sum += (item.get('配管長', 0) or 0) + (item.get('配管長2', 0) or 0)
                msp.add_line((first_x, y_count - y_count - 45000000), (first_x, y_count - y_count - 61000000))
                msp.add_line((second_x, y_count - y_count - 45000000), (second_x, y_count - y_count - 61000000))
                msp.add_line((second_x - 10000000, y_count - y_count - 45000000), (second_x - 10000000, y_count - y_count - 61000000))
                msp.add_line((first_x, y_count - y_count - 61000000), (second_x, y_count - y_count - 61000000))
                cn.flat_line1(msp, item, offset_x=x_count-16050000, offset_y=y_count - y_count - 61000000)
                cn.flat_line2(msp, item, offset_x=x_count-5000000, offset_y=y_count - y_count - 61000000)
                if item['水平距離'] == item['配管長']:
                    pass
                else:
                    cn.haikann_line1(msp, item, offset_x=x_count-25100000, offset_y=y_count - y_count - 54500000)                
                if item['水平距離2'] == item['配管長2']:
                    pass
                else:
                    cn.haikann_line2(msp, item, offset_x=x_count-12600000, offset_y=y_count - y_count - 54500000)
    elif (
        'GX形ダクタイル鋳鉄' in name
        and '曲管' in name
        and '22.5°' in name
        and '（V）' in name
        and down_22
    ):
        if item['方向'] == 'AA':
            cn.gx_vb_ryouuke_bend_22_down_ed_aa(msp, offset_x=x_count, offset_y=y_count)
            x_count += 37510000
            y_count += - 7460000
    elif (
        'GX形ダクタイル鋳鉄' in name
        and '曲管' in name
        and '22.5°' in name
        and '（V）' in name
        and up_22
    ):
        if item['方向'] == 'AA':
            cn.gx_vb_ryouuke_bend_22_up_ed_aa(msp, offset_x=x_count, offset_y=y_count)
            x_count += 37510000
            y_count += 7460000
    elif (
        'GX形ダクタイル鋳鉄' in name
        and '曲管' in name
        and '11.25°' in name
        and '（V）' in name
        and down_11
    ):
        if item['方向'] == 'AA':
            cn.gx_vb_ryouuke_bend_11_down_ed_aa(msp, offset_x=x_count, offset_y=y_count)
            x_count += 38630000
            y_count += - 3810000
    elif (
        'GX形ダクタイル鋳鉄' in name
        and '曲管' in name
        and '11.25°' in name
        and '（V）' in name
        and up_11
    ):
        if item['方向'] == 'AA':
            cn.gx_vb_ryouuke_bend_11_up_ed_aa(msp, offset_x=x_count, offset_y=y_count)
            x_count += 38630000
            y_count += 3810000

    elif (
        'GX形' in name
        and ('45' in name or '22' in name or '11' in name)
    ):
        if item['方向'] == 'AA':
            if today_count == 0:
                cn.gx_hb_ryouuke_bend_aa(msp, offset_x=x_count, offset_y=y_count)
                x_count += 39000000
                y_offsetcount += 1
                today_count += 1
                today_xbase += 39000000
            else:
                cn.choose_zairyou_no(msp, item, offset_x=x_count + 19500000, offset_y=y_count + 15000000)
                cn.name_sokutenn(msp, item, route_name, offset_x=x_count - 90000000, offset_y=y_count - y_count + 13500000 * y_offsetcount)
                cn.gx_hb_ryouuke_bend_aa(msp, offset_x=x_count, offset_y=y_count)
                first_x = x_count
                x_count += 39000000
                y_offsetcount += 1
                today_count += 1
                second_x = x_count
                flat_sum += (item.get('水平距離', 0) or 0) + (item.get('水平距離2', 0) or 0)
                haikann_sum += (item.get('配管長', 0) or 0) + (item.get('配管長2', 0) or 0)
                msp.add_line((first_x, y_count - y_count - 45000000), (first_x, y_count - y_count - 61000000))
                msp.add_line((second_x, y_count - y_count - 45000000), (second_x, y_count - y_count - 61000000))
                msp.add_line((second_x - 19500000, y_count - y_count - 45000000), (second_x - 19500000, y_count - y_count - 61000000))
                msp.add_line((first_x, y_count - y_count - 61000000), (second_x, y_count - y_count - 61000000))
                cn.flat_line1(msp, item, offset_x=x_count-29500000, offset_y=y_count - y_count - 61000000)
                cn.flat_line2(msp, item, offset_x=x_count-9750000, offset_y=y_count - y_count - 61000000)

        elif item['方向'] == 'AB':
            if today_count == 0:
                cn.gx_hb_katauke_bend_ab(msp, offset_x=x_count, offset_y=y_count)
                x_count += 27500000
                today_count += 1
                today_xbase += 27500000
            else:
                cn.choose_zairyou_no(msp, item, offset_x=x_count + 13750000, offset_y=y_count + 15000000)
                cn.name_sokutenn(msp, item, route_name, offset_x=x_count - 90000000, offset_y=y_count - y_count + 13500000 * y_offsetcount)
                cn.gx_hb_katauke_bend_ab(msp, offset_x=x_count, offset_y=y_count)
                first_x = x_count
                x_count += 27500000
                y_offsetcount += 1
                today_count += 1
                second_x = x_count
                flat_sum += (item.get('水平距離', 0) or 0) + (item.get('水平距離2', 0) or 0)
                haikann_sum += (item.get('配管長', 0) or 0) + (item.get('配管長2', 0) or 0)
                msp.add_line((first_x, y_count - y_count - 45000000), (first_x, y_count - y_count - 61000000))
                msp.add_line((second_x, y_count - y_count - 45000000), (second_x, y_count - y_count - 61000000))
                msp.add_line((second_x - 19500000, y_count - y_count - 45000000), (second_x - 19500000, y_count - y_count - 61000000))
                msp.add_line((first_x, y_count - y_count - 61000000), (second_x, y_count - y_count - 61000000))
                cn.flat_line1(msp, item, offset_x=x_count-21900000, offset_y=y_count - y_count - 61000000)
                cn.flat_line2(msp, item, offset_x=x_count-9750000, offset_y=y_count - y_count - 61000000)
        elif item['方向'] == 'BA':
            if today_count == 0:
                cn.gx_hb_katauke_bend_ba(msp, offset_x=x_count, offset_y=y_count)
                x_count += 27500000
                today_count += 1
                today_xbase += 27500000
            else:
                cn.choose_zairyou_no(msp, item, offset_x=x_count + 13750000, offset_y=y_count + 15000000)
                cn.name_sokutenn(msp, item, route_name, offset_x=x_count - 90000000, offset_y=y_count - y_count + 13500000 * y_offsetcount)
                cn.gx_hb_katauke_bend_ba(msp, offset_x=x_count, offset_y=y_count)
                first_x = x_count
                x_count += 27500000
                y_offsetcount += 1
                today_count += 1
                second_x = x_count
                flat_sum += (item.get('水平距離', 0) or 0) + (item.get('水平距離2', 0) or 0)
                haikann_sum += (item.get('配管長', 0) or 0) + (item.get('配管長2', 0) or 0)
                msp.add_line((first_x, y_count - y_count - 45000000), (first_x, y_count - y_count - 61000000))
                msp.add_line((second_x, y_count - y_count - 45000000), (second_x, y_count - y_count - 61000000))
                msp.add_line((second_x - 8000000, y_count - y_count - 45000000), (second_x - 8000000, y_count - y_count - 61000000))
                msp.add_line((first_x, y_count - y_count - 61000000), (second_x, y_count - y_count - 61000000))
                cn.flat_line1(msp, item, offset_x=x_count-17800000, offset_y=y_count - y_count - 61000000)
                cn.flat_line2(msp, item, offset_x=x_count-2400000, offset_y=y_count - y_count - 61000000)

    elif (
        'EF' in name
        and '直管' in name
        and '（甲切管）' in name
        and '(V)' in name
        and down_90
    ):
        if item['方向'] == 'BA':
            if today_count == 0:
                cn.pe_vb_tyokkann_kou_90_down_ba(msp, item, offset_x=x_count, offset_y=y_count)
                y_count += -5000000 - (item['配管長'] * 10000 / 2)
                msp.add_text("甲" + item['配管長'], dxfattribs={
                'insert': (x_count + 12000000, y_count - 8000000),
                'height': 6000000,
                })
                today_count += 1
            else:
                cn.choose_zairyou_no(msp, item, offset_x=x_count + 5000000 + (item['配管長'] * 10000 / 2), offset_y=y_count)                
                cn.pe_vb_tyokkann_kou_90_down_ba(msp, item, offset_x=x_count, offset_y=y_count)
                msp.add_text(f"甲{item['配管長']}", dxfattribs={
                'insert': (x_count - 25000000, y_count - (item['配管長'] * 10000 / 2)),
                'height': 6000000,
                })
                y_count += -5000000 - (item['配管長'] * 10000 / 2)
                today_count += 1
                flat_sum += (item.get('水平距離', 0) or 0) + (item.get('水平距離2', 0) or 0)
                haikann_sum += (item.get('配管長', 0) or 0) + (item.get('配管長2', 0) or 0)
        elif item['方向'] == 'AB':
            if today_count == 0:
                cn.pe_vb_tyokkann_kou_90_down_ab(msp, item, offset_x=x_count, offset_y=y_count)
                y_count += -5000000 - (item['配管長'] * 10000 / 2)
                msp.add_text("甲" + item['配管長'], dxfattribs={
                'insert': (x_count + 12000000, y_count - 8000000),
                'height': 6000000,
                })
                today_count += 1
            else:
                cn.choose_zairyou_no(msp, item, offset_x=x_count + 5000000 + (item['配管長'] * 10000 / 2), offset_y=y_count)                
                cn.pe_vb_tyokkann_kou_90_down_ab(msp, item, offset_x=x_count, offset_y=y_count)
                msp.add_text(f"甲{item['配管長']}", dxfattribs={
                'insert': (x_count - 25000000, y_count - (item['配管長'] * 10000 / 2)),
                'height': 6000000,
                })
                y_count += -5000000 - (item['配管長'] * 10000 / 2)
                today_count += 1
                flat_sum += (item.get('水平距離', 0) or 0) + (item.get('水平距離2', 0) or 0)
                haikann_sum += (item.get('配管長', 0) or 0) + (item.get('配管長2', 0) or 0)
    elif (
        'EF' in name
        and '直管' in name
        and '（甲切管）' in name
        and '(V)' in name
        and up_90
    ):
        if item['方向'] == 'BA':
            if today_count == 0:
                cn.pe_vb_tyokkann_kou_90_up_ba(msp, item, offset_x=x_count, offset_y=y_count)
                y_count += 5000000 + (item['配管長'] * 10000 / 2)
                msp.add_text("甲" + item['配管長'], dxfattribs={
                'insert': (x_count + 12000000, y_count - 8000000),
                'height': 6000000,
                })
                today_count += 1
            else:
                cn.choose_zairyou_no(msp, item, offset_x=x_count + 10000000, offset_y=y_count + (item['配管長'] * 10000 / 2))                
                cn.pe_vb_tyokkann_kou_90_up_ba(msp, item, offset_x=x_count, offset_y=y_count)
                msp.add_text(f"甲{item['配管長']}", dxfattribs={
                'insert': (x_count + 15000000, y_count + (item['配管長'] * 10000 / 2)),
                'height': 6000000,
                })
                y_count += 5000000 + (item['配管長'] * 10000 / 2)
                today_count += 1
                flat_sum += (item.get('水平距離', 0) or 0) + (item.get('水平距離2', 0) or 0)
                haikann_sum += (item.get('配管長', 0) or 0) + (item.get('配管長2', 0) or 0)
        elif item['方向'] == 'AB':
            if today_count == 0:
                cn.pe_vb_tyokkann_kou_90_up_ab(msp, item, offset_x=x_count, offset_y=y_count)
                y_count += 5000000 + (item['配管長'] * 10000 / 2)
                msp.add_text("甲" + item['配管長'], dxfattribs={
                'insert': (x_count + 12000000, y_count - 8000000),
                'height': 6000000,
                })
                today_count += 1
            else:
                cn.choose_zairyou_no(msp, item, offset_x=x_count + 10000000, offset_y=y_count + (item['配管長'] * 10000 / 2))                
                cn.pe_vb_tyokkann_kou_90_up_ab(msp, item, offset_x=x_count, offset_y=y_count)
                msp.add_text(f"甲{item['配管長']}", dxfattribs={
                'insert': (x_count + 15000000, y_count + (item['配管長'] * 10000 / 2)),
                'height': 6000000,
                })
                y_count += 5000000 + (item['配管長'] * 10000 / 2)
                today_count += 1
                flat_sum += (item.get('水平距離', 0) or 0) + (item.get('水平距離2', 0) or 0)
                haikann_sum += (item.get('配管長', 0) or 0) + (item.get('配管長2', 0) or 0)
    elif (
        'EF' in name
        and '直管' in name
        and '（甲切管）' in name
        and '(V)' in name
        and down_45
    ):
        if item['方向'] == 'BA':
            if today_count == 0:
                cn.pe_vb_tyokkann_kou_45_down_ba(msp, item, offset_x=x_count, offset_y=y_count)
                x_count += 3500000 + (item['水平距離'] * 10000 / 2)
                y_count += -3500000 - (item['水平距離'] * 10000 / 2)
                msp.add_text(f"甲{item['配管長']}", dxfattribs={
                'insert': (x_count - 20000000, y_count - 6000000 - (item['配管長'] * 10000 / 2)),
                'height': 6000000,
                })
                today_xbase += 3500000 + (item['水平距離'] * 10000 / 2)
                today_count += 1
            else:
                cn.choose_zairyou_no(msp, item, offset_x=x_count + 10000000, offset_y=y_count + (item['配管長'] * 10000 / 2))                
                cn.pe_vb_tyokkann_kou_45_down_ba(msp, item, offset_x=x_count, offset_y=y_count)
                msp.add_text(f"甲{item['配管長']}", dxfattribs={
                'insert': (x_count - 20000000, y_count - 6000000 - (item['配管長'] * 10000 / 2)),
                'height': 6000000,
                })
                first_x = x_count
                x_count += 3500000 + (item['水平距離'] * 10000 / 2)
                y_count += -3500000 - (item['水平距離'] * 10000 / 2)
                today_count += 1
                second_x = x_count
                flat_sum += (item.get('水平距離', 0) or 0) + (item.get('水平距離2', 0) or 0)
                haikann_sum += (item.get('配管長', 0) or 0) + (item.get('配管長2', 0) or 0)
                msp.add_line((first_x, y_count - y_count - 45000000), (first_x, y_count - y_count - 61000000))
                msp.add_line((second_x, y_count - y_count - 45000000), (second_x, y_count - y_count - 61000000))
                msp.add_line((first_x, y_count - y_count - 61000000), (second_x, y_count - y_count - 61000000))
                cn.flat_line1(msp, item, offset_x=x_count + 1800000 - (3500000 + (item['水平距離'] * 10000 / 4)), offset_y=y_count - y_count - 48000000)
                if item['水平距離'] == item['配管長']:
                    pass
                else:
                    cn.haikann_line1(msp, item, offset_x=x_count -7400000 - (3500000 + (item['水平距離'] * 10000 / 4)), offset_y=y_count - y_count - 42000000)                   
        elif item['方向'] == 'AB':
            if today_count == 0:
                cn.pe_vb_tyokkann_kou_45_down_ab(msp, item, offset_x=x_count, offset_y=y_count)
                x_count += 3500000 + (item['水平距離'] * 10000 / 2)
                y_count += -3500000 - (item['水平距離'] * 10000 / 2)
                msp.add_text(f"甲{item['配管長']}", dxfattribs={
                'insert': (x_count - 20000000, y_count - 6000000 - (item['配管長'] * 10000 / 2)),
                'height': 6000000,
                })
                today_xbase += 3500000 + (item['水平距離'] * 10000 / 2)
                today_count += 1
            else:
                cn.choose_zairyou_no(msp, item, offset_x=x_count + 10000000, offset_y=y_count + (item['配管長'] * 10000 / 2))                
                cn.pe_vb_tyokkann_kou_45_down_ab(msp, item, offset_x=x_count, offset_y=y_count)
                msp.add_text(f"甲{item['配管長']}", dxfattribs={
                'insert': (x_count - 20000000, y_count - 6000000 - (item['配管長'] * 10000 / 2)),
                'height': 6000000,
                })
                first_x = x_count
                x_count += 3500000 + (item['水平距離'] * 10000 / 2)
                y_count += -3500000 - (item['水平距離'] * 10000 / 2)
                today_count += 1
                second_x = x_count
                flat_sum += (item.get('水平距離', 0) or 0) + (item.get('水平距離2', 0) or 0)
                haikann_sum += (item.get('配管長', 0) or 0) + (item.get('配管長2', 0) or 0)
                msp.add_line((first_x, y_count - y_count - 45000000), (first_x, y_count - y_count - 61000000))
                msp.add_line((second_x, y_count - y_count - 45000000), (second_x, y_count - y_count - 61000000))
                msp.add_line((first_x, y_count - y_count - 61000000), (second_x, y_count - y_count - 61000000))
                cn.flat_line1(msp, item, offset_x=x_count + 1800000 - (3500000 + (item['水平距離'] * 10000 / 4)), offset_y=y_count - y_count - 48000000)
                if item['水平距離'] == item['配管長']:
                    pass
                else:
                    cn.haikann_line1(msp, item, offset_x=x_count -7400000 - (3500000 + (item['水平距離'] * 10000 / 4)), offset_y=y_count - y_count - 42000000)     
    elif (
        'EF' in name
        and '直管' in name
        and '（甲切管）' in name
        and '(V)' in name
        and up_45
    ):
        if item['方向'] == 'BA':
            if today_count == 0:
                cn.pe_vb_tyokkann_kou_45_up_ba(msp, item, offset_x=x_count, offset_y=y_count)
                x_count += 3500000 + (item['水平距離'] * 10000 / 2)
                y_count += 3500000 + (item['水平距離'] * 10000 / 2)
                msp.add_text(f"甲{item['配管長']}", dxfattribs={
                'insert': (x_count + 5000000, y_count - 3000000 - (item['配管長'] * 10000 / 2)),
                'height': 6000000,
                })
                today_xbase += 3500000 + (item['水平距離'] * 10000 / 2)
                today_count += 1
            else:
                cn.choose_zairyou_no(msp, item, offset_x=x_count, offset_y=y_count + 12000000 + (item['配管長'] * 10000 / 2))
                cn.pe_vb_tyokkann_kou_45_up_ba(msp, item, offset_x=x_count, offset_y=y_count)
                msp.add_text(f"甲{item['配管長']}", dxfattribs={
                'insert': (x_count + 5000000, y_count - 3000000- (item['配管長'] * 10000 / 2)),
                'height': 6000000,
                })
                first_x = x_count
                x_count += 3500000 + (item['水平距離'] * 10000 / 2)
                y_count += 3500000 + (item['水平距離'] * 10000 / 2)
                today_count += 1
                second_x = x_count
                flat_sum += (item.get('水平距離', 0) or 0) + (item.get('水平距離2', 0) or 0)
                haikann_sum += (item.get('配管長', 0) or 0) + (item.get('配管長2', 0) or 0)
                msp.add_line((first_x, y_count - y_count - 45000000), (first_x, y_count - y_count - 61000000))
                msp.add_line((second_x, y_count - y_count - 45000000), (second_x, y_count - y_count - 61000000))
                msp.add_line((first_x, y_count - y_count - 61000000), (second_x, y_count - y_count - 61000000))
                cn.flat_line1(msp, item, offset_x=x_count + 1800000 - (3500000 + (item['水平距離'] * 10000 / 4)), offset_y=y_count - y_count - 48000000)
                if item['水平距離'] == item['配管長']:
                    pass
                else:
                    cn.haikann_line1(msp, item, offset_x=x_count -7400000 - (3500000 + (item['水平距離'] * 10000 / 4)), offset_y=y_count - y_count - 42000000)    
        elif item['方向'] == 'AB':
            if today_count == 0:
                cn.pe_vb_tyokkann_kou_45_up_ab(msp, item, offset_x=x_count, offset_y=y_count)
                x_count += 3500000 + (item['水平距離'] * 10000 / 2)
                y_count += 3500000 + (item['水平距離'] * 10000 / 2)
                msp.add_text(f"甲{item['配管長']}", dxfattribs={
                'insert': (x_count + 5000000, y_count - 3000000 - (item['配管長'] * 10000 / 2)),
                'height': 6000000,
                })
                today_xbase += 3500000 + (item['水平距離'] * 10000 / 2)
                today_count += 1
            else:
                cn.choose_zairyou_no(msp, item, offset_x=x_count, offset_y=y_count + 12000000 + (item['配管長'] * 10000 / 2))
                cn.pe_vb_tyokkann_kou_45_up_ab(msp, item, offset_x=x_count, offset_y=y_count)
                msp.add_text(f"甲{item['配管長']}", dxfattribs={
                'insert': (x_count + 5000000, y_count - 3000000- (item['配管長'] * 10000 / 2)),
                'height': 6000000,
                })
                first_x = x_count
                x_count += 3500000 + (item['水平距離'] * 10000 / 2)
                y_count += 3500000 + (item['水平距離'] * 10000 / 2)
                today_count += 1
                second_x = x_count
                flat_sum += (item.get('水平距離', 0) or 0) + (item.get('水平距離2', 0) or 0)
                haikann_sum += (item.get('配管長', 0) or 0) + (item.get('配管長2', 0) or 0)
                msp.add_line((first_x, y_count - y_count - 45000000), (first_x, y_count - y_count - 61000000))
                msp.add_line((second_x, y_count - y_count - 45000000), (second_x, y_count - y_count - 61000000))
                msp.add_line((first_x, y_count - y_count - 61000000), (second_x, y_count - y_count - 61000000))
                cn.flat_line1(msp, item, offset_x=x_count + 1800000 - (3500000 + (item['水平距離'] * 10000 / 4)), offset_y=y_count - y_count - 48000000)
                if item['水平距離'] == item['配管長']:
                    pass
                else:
                    cn.haikann_line1(msp, item, offset_x=x_count -7400000 - (3500000 + (item['水平距離'] * 10000 / 4)), offset_y=y_count - y_count - 42000000)

    elif (
        'EF' in name
        and '直管' in name
        and '（甲切管）' in name
        and '(V)' in name
        and down_22
    ):
        if item['方向'] == 'BA':
            if today_count == 0:
                cn.pe_vb_tyokkann_kou_22_down_ba(msp, item, offset_x=x_count, offset_y=y_count)
                x_count += 4620000 + item['水平距離'] * 10000 / 2
                y_count += - 1910000 - ((item['配管長'] * 10000 / 2) ** 2 - (item['水平距離'] * 10000 / 2) ** 2) ** 0.5
                msp.add_text(f"甲{item['配管長']}", dxfattribs={
                'insert': (x_count - (4620000 + item['水平距離'] * 10000 / 4), y_count - 12000000 - 1910000 - ((item['配管長'] * 10000 / 2) ** 2 - (item['水平距離'] * 10000 / 2) ** 2) ** 0.5),
                'height': 6000000,
                })
                today_xbase += 4620000 + item['水平距離'] * 10000 / 2
                today_count += 1
            else:
                cn.choose_zairyou_no(msp, item, offset_x=x_count + 4620000 + item['水平距離'] * 10000 / 4, offset_y=y_count + 15000000 - 1910000 - ((item['配管長'] * 10000 / 2) ** 2 - (item['水平距離'] * 10000 / 2) ** 2) ** 0.5)
                cn.pe_vb_tyokkann_kou_22_down_ba(msp, item, offset_x=x_count, offset_y=y_count)
                msp.add_text(f"甲{item['配管長']}", dxfattribs={
                'insert': (x_count - (4620000 + item['水平距離'] * 10000 / 4), y_count - 12000000 - 1910000 - ((item['配管長'] * 10000 / 2) ** 2 - (item['水平距離'] * 10000 / 2) ** 2) ** 0.5),
                'height': 6000000,
                })
                first_x = x_count
                x_count += 4620000 + item['水平距離'] * 10000 / 2
                y_count += - 1910000 - ((item['配管長'] * 10000 / 2) ** 2 - (item['水平距離'] * 10000 / 2) ** 2) ** 0.5
                today_count += 1
                second_x = x_count
                flat_sum += (item.get('水平距離', 0) or 0) + (item.get('水平距離2', 0) or 0)
                haikann_sum += (item.get('配管長', 0) or 0) + (item.get('配管長2', 0) or 0)
                msp.add_line((first_x, y_count - y_count - 45000000), (first_x, y_count - y_count - 61000000))
                msp.add_line((second_x, y_count - y_count - 45000000), (second_x, y_count - y_count - 61000000))
                msp.add_line((first_x, y_count - y_count - 61000000), (second_x, y_count - y_count - 61000000))
                cn.flat_line1(msp, item, offset_x=x_count + 2300000 - (4620000 + item['水平距離'] * 10000 / 4), offset_y=y_count - y_count - 48000000)
                if item['水平距離'] == item['配管長']:
                    pass
                else:
                    cn.haikann_line1(msp, item, offset_x=x_count -6900000 - (4620000 + item['水平距離'] * 10000 / 4), offset_y=y_count - y_count - 42000000)
        elif item['方向'] == 'AB':
            if today_count == 0:
                cn.pe_vb_tyokkann_kou_22_down_ab(msp, item, offset_x=x_count, offset_y=y_count)
                x_count += 4620000 + item['水平距離'] * 10000 / 2
                y_count += - 1910000 - ((item['配管長'] * 10000 / 2) ** 2 - (item['水平距離'] * 10000 / 2) ** 2) ** 0.5
                msp.add_text(f"甲{item['配管長']}", dxfattribs={
                'insert': (x_count - (4620000 + item['水平距離'] * 10000 / 4), y_count - 12000000 - 1910000 - ((item['配管長'] * 10000 / 2) ** 2 - (item['水平距離'] * 10000 / 2) ** 2) ** 0.5),
                'height': 6000000,
                })
                today_xbase += 4620000 + item['水平距離'] * 10000 / 2
                today_count += 1
            else:
                cn.choose_zairyou_no(msp, item, offset_x=x_count + 4620000 + item['水平距離'] * 10000 / 4, offset_y=y_count + 15000000 - 1910000 - ((item['配管長'] * 10000 / 2) ** 2 - (item['水平距離'] * 10000 / 2) ** 2) ** 0.5)
                cn.pe_vb_tyokkann_kou_22_down_ab(msp, item, offset_x=x_count, offset_y=y_count)
                msp.add_text(f"甲{item['配管長']}", dxfattribs={
                'insert': (x_count - (4620000 + item['水平距離'] * 10000 / 4), y_count - 12000000 - 1910000 - ((item['配管長'] * 10000 / 2) ** 2 - (item['水平距離'] * 10000 / 2) ** 2) ** 0.5),
                'height': 6000000,
                })
                first_x = x_count
                x_count += 4620000 + item['水平距離'] * 10000 / 2
                y_count += - 1910000 - ((item['配管長'] * 10000 / 2) ** 2 - (item['水平距離'] * 10000 / 2) ** 2) ** 0.5
                today_count += 1
                second_x = x_count
                flat_sum += (item.get('水平距離', 0) or 0) + (item.get('水平距離2', 0) or 0)
                haikann_sum += (item.get('配管長', 0) or 0) + (item.get('配管長2', 0) or 0)
                msp.add_line((first_x, y_count - y_count - 45000000), (first_x, y_count - y_count - 61000000))
                msp.add_line((second_x, y_count - y_count - 45000000), (second_x, y_count - y_count - 61000000))
                msp.add_line((first_x, y_count - y_count - 61000000), (second_x, y_count - y_count - 61000000))
                cn.flat_line1(msp, item, offset_x=x_count + 2300000 - (4620000 + item['水平距離'] * 10000 / 4), offset_y=y_count - y_count - 48000000)
                if item['水平距離'] == item['配管長']:
                    pass
                else:
                    cn.haikann_line1(msp, item, offset_x=x_count -6900000 - (4620000 + item['水平距離'] * 10000 / 4), offset_y=y_count - y_count - 42000000)
    elif (
        'EF' in name
        and '直管' in name
        and '（甲切管）' in name
        and '(V)' in name
        and up_22
    ):
        if item['方向'] == 'BA':
            if today_count == 0:
                cn.pe_vb_tyokkann_kou_22_up_ba(msp, item, offset_x=x_count, offset_y=y_count)
                x_count += 4620000 + item['水平距離'] * 10000 / 2
                y_count += + 1910000 + ((item['配管長'] * 10000 / 2) ** 2 - (item['水平距離'] * 10000 / 2) ** 2) ** 0.5
                msp.add_text(f"甲{item['配管長']}", dxfattribs={
                'insert': (x_count - (4620000 + item['水平距離'] * 10000 / 4), y_count - 16000000 + 1910000 + ((item['配管長'] * 10000 / 2) ** 2 - (item['水平距離'] * 10000 / 2) ** 2) ** 0.5),
                'height': 6000000,
                })
                today_xbase += 4620000 + item['水平距離'] * 10000 / 2
                today_count += 1
            else:
                cn.choose_zairyou_no(msp, item, offset_x=x_count + 4620000 + item['水平距離'] * 10000 / 4, offset_y=y_count + 15000000 + 1910000 + ((item['配管長'] * 10000 / 2) ** 2 - (item['水平距離'] * 10000 / 2) ** 2) ** 0.5)
                cn.pe_vb_tyokkann_kou_22_up_ba(msp, item, offset_x=x_count, offset_y=y_count)
                msp.add_text(f"甲{item['配管長']}", dxfattribs={
                'insert': (x_count - (4620000 + item['水平距離'] * 10000 / 4), y_count - 16000000 + 1910000 + ((item['配管長'] * 10000 / 2) ** 2 - (item['水平距離'] * 10000 / 2) ** 2) ** 0.5),
                'height': 6000000,
                })
                first_x = x_count
                x_count += 4620000 + item['水平距離'] * 10000 / 2
                y_count += 1910000 + ((item['配管長'] * 10000 / 2) ** 2 - (item['水平距離'] * 10000 / 2) ** 2) ** 0.5
                today_count += 1
                second_x = x_count
                flat_sum += (item.get('水平距離', 0) or 0) + (item.get('水平距離2', 0) or 0)
                haikann_sum += (item.get('配管長', 0) or 0) + (item.get('配管長2', 0) or 0)
                msp.add_line((first_x, y_count - y_count - 45000000), (first_x, y_count - y_count - 61000000))
                msp.add_line((second_x, y_count - y_count - 45000000), (second_x, y_count - y_count - 61000000))
                msp.add_line((first_x, y_count - y_count - 61000000), (second_x, y_count - y_count - 61000000))
                cn.flat_line1(msp, item, offset_x=x_count + 2300000 - (4620000 + item['水平距離'] * 10000 / 4), offset_y=y_count - y_count - 48000000)
                if item['水平距離'] == item['配管長']:
                    pass
                else:
                    cn.haikann_line1(msp, item, offset_x=x_count -6900000 - (4620000 + item['水平距離'] * 10000 / 4), offset_y=y_count - y_count - 42000000)
        elif item['方向'] == 'AB':
            if today_count == 0:
                cn.pe_vb_tyokkann_kou_22_up_ab(msp, item, offset_x=x_count, offset_y=y_count)
                x_count += 4620000 + item['水平距離'] * 10000 / 2
                y_count += + 1910000 + ((item['配管長'] * 10000 / 2) ** 2 - (item['水平距離'] * 10000 / 2) ** 2) ** 0.5
                msp.add_text(f"甲{item['配管長']}", dxfattribs={
                'insert': (x_count - (4620000 + item['水平距離'] * 10000 / 4), y_count - 16000000 + 1910000 + ((item['配管長'] * 10000 / 2) ** 2 - (item['水平距離'] * 10000 / 2) ** 2) ** 0.5),
                'height': 6000000,
                })
                today_xbase += 4620000 + item['水平距離'] * 10000 / 2
                today_count += 1
            else:
                cn.choose_zairyou_no(msp, item, offset_x=x_count + 4620000 + item['水平距離'] * 10000 / 4, offset_y=y_count + 15000000 + 1910000 + ((item['配管長'] * 10000 / 2) ** 2 - (item['水平距離'] * 10000 / 2) ** 2) ** 0.5)
                cn.pe_vb_tyokkann_kou_22_up_ab(msp, item, offset_x=x_count, offset_y=y_count)
                msp.add_text(f"甲{item['配管長']}", dxfattribs={
                'insert': (x_count - (4620000 + item['水平距離'] * 10000 / 4), y_count - 16000000 + 1910000 + ((item['配管長'] * 10000 / 2) ** 2 - (item['水平距離'] * 10000 / 2) ** 2) ** 0.5),
                'height': 6000000,
                })
                first_x = x_count
                x_count += 4620000 + item['水平距離'] * 10000 / 2
                y_count += 1910000 + ((item['配管長'] * 10000 / 2) ** 2 - (item['水平距離'] * 10000 / 2) ** 2) ** 0.5
                today_count += 1
                second_x = x_count
                flat_sum += (item.get('水平距離', 0) or 0) + (item.get('水平距離2', 0) or 0)
                haikann_sum += (item.get('配管長', 0) or 0) + (item.get('配管長2', 0) or 0)
                msp.add_line((first_x, y_count - y_count - 45000000), (first_x, y_count - y_count - 61000000))
                msp.add_line((second_x, y_count - y_count - 45000000), (second_x, y_count - y_count - 61000000))
                msp.add_line((first_x, y_count - y_count - 61000000), (second_x, y_count - y_count - 61000000))
                cn.flat_line1(msp, item, offset_x=x_count + 2300000 - (4620000 + item['水平距離'] * 10000 / 4), offset_y=y_count - y_count - 48000000)
                if item['水平距離'] == item['配管長']:
                    pass
                else:
                    cn.haikann_line1(msp, item, offset_x=x_count -6900000 - (4620000 + item['水平距離'] * 10000 / 4), offset_y=y_count - y_count - 42000000)
    elif (
        'EF' in name
        and '直管' in name
        and '（甲切管）' in name
        and '(V)' in name
        and down_11
    ):
        if item['方向'] == 'BA':
            if today_count == 0:
                cn.pe_vb_tyokkann_kou_11_down_ba(msp, item, offset_x=x_count, offset_y=y_count)
                x_count += 4900000 + item['水平距離'] * 10000 / 2
                y_count += - 980000 - ((item['配管長'] * 10000 / 2) ** 2 - (item['水平距離'] * 10000 / 2) ** 2) ** 0.5
                msp.add_text(f"甲{item['配管長']}", dxfattribs={
                'insert': (x_count - (4900000 + item['水平距離'] * 10000 / 4), y_count - 12000000 - 980000 - ((item['配管長'] * 10000 / 2) ** 2 - (item['水平距離'] * 10000 / 2) ** 2) ** 0.5),
                'height': 6000000,
                })
                today_xbase += 4900000 + item['水平距離'] * 10000 / 2
                today_count += 1
            else:
                cn.choose_zairyou_no(msp, item, offset_x=x_count + 4900000 + item['水平距離'] * 10000 / 4, offset_y=y_count + 15000000 - 980000 - ((item['配管長'] * 10000 / 2) ** 2 - (item['水平距離'] * 10000 / 2) ** 2) ** 0.5)
                cn.pe_vb_tyokkann_kou_11_down_ba(msp, item, offset_x=x_count, offset_y=y_count)
                msp.add_text(f"甲{item['配管長']}", dxfattribs={
                'insert': (x_count - (4900000 + item['水平距離'] * 10000 / 4), y_count - 12000000 - 980000 - ((item['配管長'] * 10000 / 2) ** 2 - (item['水平距離'] * 10000 / 2) ** 2) ** 0.5),
                'height': 6000000,
                })
                first_x = x_count
                x_count += 4900000 + item['水平距離'] * 10000 / 2
                y_count += - 980000 - ((item['配管長'] * 10000 / 2) ** 2 - (item['水平距離'] * 10000 / 2) ** 2) ** 0.5
                today_count += 1
                second_x = x_count
                flat_sum += (item.get('水平距離', 0) or 0) + (item.get('水平距離2', 0) or 0)
                haikann_sum += (item.get('配管長', 0) or 0) + (item.get('配管長2', 0) or 0)
                msp.add_line((first_x, y_count - y_count - 45000000), (first_x, y_count - y_count - 61000000))
                msp.add_line((second_x, y_count - y_count - 45000000), (second_x, y_count - y_count - 61000000))
                msp.add_line((first_x, y_count - y_count - 61000000), (second_x, y_count - y_count - 61000000))
                cn.flat_line1(msp, item, offset_x=x_count + 2300000 - (4900000 + item['水平距離'] * 10000 / 4), offset_y=y_count - y_count - 48000000)
                if item['水平距離'] == item['配管長']:
                    pass
                else:
                    cn.haikann_line1(msp, item, offset_x=x_count -6900000 - (4900000 + item['水平距離'] * 10000 / 4), offset_y=y_count - y_count - 42000000)
        elif item['方向'] == 'AB':
            if today_count == 0:
                cn.pe_vb_tyokkann_kou_11_down_ab(msp, item, offset_x=x_count, offset_y=y_count)
                x_count += 4900000 + item['水平距離'] * 10000 / 2
                y_count += - 980000 - ((item['配管長'] * 10000 / 2) ** 2 - (item['水平距離'] * 10000 / 2) ** 2) ** 0.5
                msp.add_text(f"甲{item['配管長']}", dxfattribs={
                'insert': (x_count - (4900000 + item['水平距離'] * 10000 / 4), y_count - 12000000 - 980000 - ((item['配管長'] * 10000 / 2) ** 2 - (item['水平距離'] * 10000 / 2) ** 2) ** 0.5),
                'height': 6000000,
                })
                today_xbase += 4900000 + item['水平距離'] * 10000 / 2
                today_count += 1
            else:
                cn.choose_zairyou_no(msp, item, offset_x=x_count + 4900000 + item['水平距離'] * 10000 / 4, offset_y=y_count + 15000000 - 980000 - ((item['配管長'] * 10000 / 2) ** 2 - (item['水平距離'] * 10000 / 2) ** 2) ** 0.5)
                cn.pe_vb_tyokkann_kou_11_down_ab(msp, item, offset_x=x_count, offset_y=y_count)
                msp.add_text(f"甲{item['配管長']}", dxfattribs={
                'insert': (x_count - (4900000 + item['水平距離'] * 10000 / 4), y_count - 12000000 - 980000 - ((item['配管長'] * 10000 / 2) ** 2 - (item['水平距離'] * 10000 / 2) ** 2) ** 0.5),
                'height': 6000000,
                })
                first_x = x_count
                x_count += 4900000 + item['水平距離'] * 10000 / 2
                y_count += - 980000 - ((item['配管長'] * 10000 / 2) ** 2 - (item['水平距離'] * 10000 / 2) ** 2) ** 0.5
                today_count += 1
                second_x = x_count
                flat_sum += (item.get('水平距離', 0) or 0) + (item.get('水平距離2', 0) or 0)
                haikann_sum += (item.get('配管長', 0) or 0) + (item.get('配管長2', 0) or 0)
                msp.add_line((first_x, y_count - y_count - 45000000), (first_x, y_count - y_count - 61000000))
                msp.add_line((second_x, y_count - y_count - 45000000), (second_x, y_count - y_count - 61000000))
                msp.add_line((first_x, y_count - y_count - 61000000), (second_x, y_count - y_count - 61000000))
                cn.flat_line1(msp, item, offset_x=x_count + 2300000 - (4900000 + item['水平距離'] * 10000 / 4), offset_y=y_count - y_count - 48000000)
                if item['水平距離'] == item['配管長']:
                    pass
                else:
                    cn.haikann_line1(msp, item, offset_x=x_count -6900000 - (4900000 + item['水平距離'] * 10000 / 4), offset_y=y_count - y_count - 42000000)

    elif (
        'EF' in name
        and '直管' in name
        and '（甲切管）' in name
        and '(V)' in name
        and up_11
    ):
        if item['方向'] == 'BA':
            if today_count == 0:
                cn.pe_vb_tyokkann_kou_11_up_ba(msp, item, offset_x=x_count, offset_y=y_count)
                x_count += 4900000 + item['水平距離'] * 10000 / 2
                y_count += + 980000 + ((item['配管長'] * 10000 / 2) ** 2 - (item['水平距離'] * 10000 / 2) ** 2) ** 0.5
                msp.add_text(f"甲{item['配管長']}", dxfattribs={
                'insert': (x_count - (4900000 + item['水平距離'] * 10000 / 4), y_count - 16000000 + 980000 + ((item['配管長'] * 10000 / 2) ** 2 - (item['水平距離'] * 10000 / 2) ** 2) ** 0.5),
                'height': 6000000,
                })
                today_xbase += 4900000 + item['水平距離'] * 10000 / 2
                today_count += 1
            else:
                cn.choose_zairyou_no(msp, item, offset_x=x_count + 4900000 + item['水平距離'] * 10000 / 4, offset_y=y_count + 15000000 + 1910000 + ((item['配管長'] * 10000 / 2) ** 2 - (item['水平距離'] * 10000 / 2) ** 2) ** 0.5)
                cn.pe_vb_tyokkann_kou_11_up_ba(msp, item, offset_x=x_count, offset_y=y_count)
                msp.add_text(f"甲{item['配管長']}", dxfattribs={
                'insert': (x_count - (4900000 + item['水平距離'] * 10000 / 4), y_count - 16000000 + 980000 + ((item['配管長'] * 10000 / 2) ** 2 - (item['水平距離'] * 10000 / 2) ** 2) ** 0.5),
                'height': 6000000,
                })
                first_x = x_count
                x_count += 4900000 + item['水平距離'] * 10000 / 2
                y_count += 980000 + ((item['配管長'] * 10000 / 2) ** 2 - (item['水平距離'] * 10000 / 2) ** 2) ** 0.5
                today_count += 1
                second_x = x_count
                flat_sum += (item.get('水平距離', 0) or 0) + (item.get('水平距離2', 0) or 0)
                haikann_sum += (item.get('配管長', 0) or 0) + (item.get('配管長2', 0) or 0)
                msp.add_line((first_x, y_count - y_count - 45000000), (first_x, y_count - y_count - 61000000))
                msp.add_line((second_x, y_count - y_count - 45000000), (second_x, y_count - y_count - 61000000))
                msp.add_line((first_x, y_count - y_count - 61000000), (second_x, y_count - y_count - 61000000))
                cn.flat_line1(msp, item, offset_x=x_count + 2300000 - (4900000 + item['水平距離'] * 10000 / 4), offset_y=y_count - y_count - 48000000)
                if item['水平距離'] == item['配管長']:
                    pass
                else:
                    cn.haikann_line1(msp, item, offset_x=x_count -6900000 - (4900000 + item['水平距離'] * 10000 / 4), offset_y=y_count - y_count - 42000000)
        elif item['方向'] == 'AB':
            if today_count == 0:
                cn.pe_vb_tyokkann_kou_11_up_ab(msp, item, offset_x=x_count, offset_y=y_count)
                x_count += 4900000 + item['水平距離'] * 10000 / 2
                y_count += + 980000 + ((item['配管長'] * 10000 / 2) ** 2 - (item['水平距離'] * 10000 / 2) ** 2) ** 0.5
                msp.add_text(f"甲{item['配管長']}", dxfattribs={
                'insert': (x_count - (4900000 + item['水平距離'] * 10000 / 4), y_count - 16000000 + 980000 + ((item['配管長'] * 10000 / 2) ** 2 - (item['水平距離'] * 10000 / 2) ** 2) ** 0.5),
                'height': 6000000,
                })
                today_xbase += 4900000 + item['水平距離'] * 10000 / 2
                today_count += 1
            else:
                cn.choose_zairyou_no(msp, item, offset_x=x_count + 4900000 + item['水平距離'] * 10000 / 4, offset_y=y_count + 15000000 + 1910000 + ((item['配管長'] * 10000 / 2) ** 2 - (item['水平距離'] * 10000 / 2) ** 2) ** 0.5)
                cn.pe_vb_tyokkann_kou_22_up_ab(msp, item, offset_x=x_count, offset_y=y_count)
                msp.add_text(f"甲{item['配管長']}", dxfattribs={
                'insert': (x_count - (4900000 + item['水平距離'] * 10000 / 4), y_count - 16000000 + 980000 + ((item['配管長'] * 10000 / 2) ** 2 - (item['水平距離'] * 10000 / 2) ** 2) ** 0.5),
                'height': 6000000,
                })
                first_x = x_count
                x_count += 4900000 + item['水平距離'] * 10000 / 2
                y_count += 980000 + ((item['配管長'] * 10000 / 2) ** 2 - (item['水平距離'] * 10000 / 2) ** 2) ** 0.5
                today_count += 1
                second_x = x_count
                flat_sum += (item.get('水平距離', 0) or 0) + (item.get('水平距離2', 0) or 0)
                haikann_sum += (item.get('配管長', 0) or 0) + (item.get('配管長2', 0) or 0)
                msp.add_line((first_x, y_count - y_count - 45000000), (first_x, y_count - y_count - 61000000))
                msp.add_line((second_x, y_count - y_count - 45000000), (second_x, y_count - y_count - 61000000))
                msp.add_line((first_x, y_count - y_count - 61000000), (second_x, y_count - y_count - 61000000))
                cn.flat_line1(msp, item, offset_x=x_count + 2300000 - (4900000 + item['水平距離'] * 10000 / 4), offset_y=y_count - y_count - 48000000)
                if item['水平距離'] == item['配管長']:
                    pass
                else:
                    cn.haikann_line1(msp, item, offset_x=x_count -6900000 - (4900000 + item['水平距離'] * 10000 / 4), offset_y=y_count - y_count - 42000000)

    elif (
        'EF' in name
        and '直管' in name
        and '（甲切管）' in name
    ):
        if item['方向'] == 'BA':
            if today_count == 0:
                cn.pe_tyokkann_kou_ba(msp, item, offset_x=x_count, offset_y=y_count)
                x_count += 5000000 + (item['配管長'] * 10000 / 2)
                today_xbase += 5000000 + (item['配管長'] * 10000 / 2)
                today_count += 1
            else:
                cn.choose_zairyou_no(msp, item, offset_x=x_count + (5000000 + (item['配管長'] * 10000 / 2)) / 2, offset_y=y_count + 15000000)
                cn.pe_tyokkann_kou_ba(msp, item, offset_x=x_count, offset_y=y_count)
                msp.add_text(f"甲{item['配管長']}", dxfattribs={
                'insert': (x_count - 10000000 + (5000000 + (item['配管長'] * 10000 / 2)) / 2, y_count - 12000000),
                'height': 6000000,
                })
                first_x = x_count
                x_count += 5000000 + (item['配管長'] * 10000 / 2)
                today_count += 1
                second_x = x_count
                flat_sum += (item.get('水平距離', 0) or 0) + (item.get('水平距離2', 0) or 0)
                haikann_sum += (item.get('配管長', 0) or 0) + (item.get('配管長2', 0) or 0)
                msp.add_line((first_x, y_count - y_count - 45000000), (first_x, y_count - y_count - 61000000))
                msp.add_line((second_x, y_count - y_count - 45000000), (second_x, y_count - y_count - 61000000))
                msp.add_line((first_x, y_count - y_count - 61000000), (second_x, y_count - y_count - 61000000))
                cn.flat_line1(msp, item, offset_x=x_count - 1450000 - (5000000 + (item['配管長'] * 10000 / 2)) / 2, offset_y=y_count - y_count - 61000000) 
        if item['方向'] == 'AB':
            if today_count == 0:
                cn.pe_tyokkann_kou_ab(msp, item, offset_x=x_count, offset_y=y_count)
                x_count += 5000000 + (item['配管長'] * 10000 / 2)
                today_xbase += 5000000 + (item['配管長'] * 10000 / 2)
                today_count += 1
            else:
                cn.choose_zairyou_no(msp, item, offset_x=x_count + (5000000 + (item['配管長'] * 10000 / 2)) / 2, offset_y=y_count + 15000000)
                cn.pe_tyokkann_kou_ab(msp, item, offset_x=x_count, offset_y=y_count)
                msp.add_text(f"甲{item['配管長']}", dxfattribs={
                'insert': (x_count - 10000000 + (5000000 + (item['配管長'] * 10000 / 2)) / 2, y_count - 12000000),
                'height': 6000000,
                })
                first_x = x_count
                x_count += 5000000 + (item['配管長'] * 10000 / 2)
                today_count += 1
                second_x = x_count
                flat_sum += (item.get('水平距離', 0) or 0) + (item.get('水平距離2', 0) or 0)
                haikann_sum += (item.get('配管長', 0) or 0) + (item.get('配管長2', 0) or 0)
                msp.add_line((first_x, y_count - y_count - 45000000), (first_x, y_count - y_count - 61000000))
                msp.add_line((second_x, y_count - y_count - 45000000), (second_x, y_count - y_count - 61000000))
                msp.add_line((first_x, y_count - y_count - 61000000), (second_x, y_count - y_count - 61000000))
                cn.flat_line1(msp, item, offset_x=x_count - 1450000 - (5000000 + (item['配管長'] * 10000 / 2)) / 2, offset_y=y_count - y_count - 61000000) 

    elif (
        'EF' in name
        and '直管' in name
    ):
        if item['方向'] == 'SA':
            cn.choose_zairyou_no(msp, item, offset_x=x_count + 15150000, offset_y=y_count+15000000)
            cn.pe_tyokkann_ba(msp, offset_x=x_count, offset_y=y_count)
            msp.add_text("直", dxfattribs={
            'insert': (x_count + 12000000, y_count - 8000000),
            'height': 6000000,
})
            first_x = x_count
            today_count += 1
            x_count += 30250000
            second_x = x_count
            flat_sum += (item.get('水平距離', 0) or 0) + (item.get('水平距離2', 0) or 0)
            haikann_sum += (item.get('配管長', 0) or 0) + (item.get('配管長2', 0) or 0)
            msp.add_line((first_x, y_count - y_count - 45000000), (first_x, y_count - y_count - 61000000))
            msp.add_line((second_x, y_count - y_count - 45000000), (second_x, y_count - y_count - 61000000))
            msp.add_line((first_x, y_count - y_count - 61000000), (second_x, y_count - y_count - 61000000))
            cn.flat_line1(msp, item, offset_x=x_count-16500000, offset_y=y_count - y_count - 61000000)

        elif item['方向'] == 'SB':
            cn.choose_zairyou_no(msp, item, offset_x=x_count + 15150000, offset_y=y_count+15000000)
            cn.pe_tyokkann_ab(msp, offset_x=x_count, offset_y=y_count)
            msp.add_text("直", dxfattribs={
            'insert': (x_count + 12000000, y_count - 8000000),
            'height': 6000000,
})
            first_x = x_count
            today_count += 1
            x_count += 30250000
            second_x = x_count
            flat_sum += (item.get('水平距離', 0) or 0) + (item.get('水平距離2', 0) or 0)
            haikann_sum += (item.get('配管長', 0) or 0) + (item.get('配管長2', 0) or 0)
            msp.add_line((first_x, y_count - y_count - 45000000), (first_x, y_count - y_count - 61000000))
            msp.add_line((second_x, y_count - y_count - 45000000), (second_x, y_count - y_count - 61000000))
            msp.add_line((first_x, y_count - y_count - 61000000), (second_x, y_count - y_count - 61000000))
            cn.flat_line1(msp, item, offset_x=x_count-16500000, offset_y=y_count - y_count - 61000000)

        if item['方向'] == 'BA':
            if today_count == 0:
                cn.pe_tyokkann_ba(msp, offset_x=x_count, offset_y=y_count)
                today_count += 1
                today_xbase += 30250000
                x_count += 30250000
            else:
                cn.choose_zairyou_no(msp, item, offset_x=x_count + 15150000, offset_y=y_count+15000000)
                cn.pe_tyokkann_ba(msp, offset_x=x_count, offset_y=y_count)
                msp.add_text("直", dxfattribs={
                'insert': (x_count + 12000000, y_count - 8000000),
                'height': 6000000,
    })
                first_x = x_count
                today_count += 1
                x_count += 30250000
                second_x = x_count
                flat_sum += (item.get('水平距離', 0) or 0) + (item.get('水平距離2', 0) or 0)
                haikann_sum += (item.get('配管長', 0) or 0) + (item.get('配管長2', 0) or 0)
                msp.add_line((first_x, y_count - y_count - 45000000), (first_x, y_count - y_count - 61000000))
                msp.add_line((second_x, y_count - y_count - 45000000), (second_x, y_count - y_count - 61000000))
                msp.add_line((first_x, y_count - y_count - 61000000), (second_x, y_count - y_count - 61000000))
                cn.flat_line1(msp, item, offset_x=x_count-16500000, offset_y=y_count - y_count - 61000000)
        elif item['方向'] == 'AB':
            if today_count == 0:
                cn.pe_tyokkann_ab(msp, offset_x=x_count, offset_y=y_count)
                today_count += 1
                today_xbase += 30250000
                x_count += 30250000
            else:
                cn.choose_zairyou_no(msp, item, offset_x=x_count + 15150000, offset_y=y_count+15000000)
                cn.pe_tyokkann_ab(msp, offset_x=x_count, offset_y=y_count)
                msp.add_text("直", dxfattribs={
                'insert': (x_count + 12000000, y_count - 8000000),
                'height': 6000000,
    })
                first_x = x_count
                today_count += 1
                x_count += 30250000
                second_x = x_count
                flat_sum += (item.get('水平距離', 0) or 0) + (item.get('水平距離2', 0) or 0)
                haikann_sum += (item.get('配管長', 0) or 0) + (item.get('配管長2', 0) or 0)
                msp.add_line((first_x, y_count - y_count - 45000000), (first_x, y_count - y_count - 61000000))
                msp.add_line((second_x, y_count - y_count - 45000000), (second_x, y_count - y_count - 61000000))
                msp.add_line((first_x, y_count - y_count - 61000000), (second_x, y_count - y_count - 61000000))
                cn.flat_line1(msp, item, offset_x=x_count-16500000, offset_y=y_count - y_count - 61000000)

    elif (
        'EF' in name
        and 'ベンド' in name
        and '90°' in name
        and '(V)' in name
        and '下がり' in name
    ):
        if item['方向'] == 'AA':
            if today_count == 0:
                cn.pe_vb_ryouuke_bend_90_down_st_aa(msp, offset_x=x_count, offset_y=y_count)
                x_count += 10000000
                y_count += -10000000
                today_count += 1
                today_xbase += 10000000
            else:
                cn.choose_zairyou_no(msp, item, offset_x=x_count + 8500000, offset_y=y_count + 15000000)
                cn.name_sokutenn(msp, item, route_name,offset_x=x_count - 90000000, offset_y=y_count + 8000000 + 13500000 * y_offsetcount)
                cn.pe_vb_ryouuke_bend_90_down_st_aa(msp, offset_x=x_count, offset_y=y_count)
                first_x = x_count
                x_count += 10000000
                y_count += -10000000
                y_offsetcount += 1
                today_count += 1
                second_x = x_count
                flat_sum += (item.get('水平距離', 0) or 0) + (item.get('水平距離2', 0) or 0)
                haikann_sum += (item.get('配管長', 0) or 0) + (item.get('配管長2', 0) or 0)
                msp.add_line((first_x, y_count - y_count - 45000000), (first_x, y_count - y_count - 61000000))
                msp.add_line((second_x, y_count - y_count - 45000000), (second_x, y_count - y_count - 61000000))
                msp.add_line((first_x, y_count - y_count - 61000000), (second_x, y_count - y_count - 61000000))
                cn.flat_line1(msp, item, offset_x=x_count-5000000, offset_y=y_count - y_count - 61000000)
        elif item['方向'] == 'BA':
            if today_count == 0:
                cn.pe_vb_katauke_bend_90_down_st_ba(msp, offset_x=x_count, offset_y=y_count)
                x_count += 8000000
                y_count += - 10000000
                today_count += 1
                today_xbase += 8000000
            else:
                cn.choose_zairyou_no(msp, item, offset_x=x_count + 4000000, offset_y=y_count + 15000000)
                cn.name_sokutenn(msp, item, route_name,offset_x=x_count - 90000000, offset_y=y_count + 8000000 + 13500000 * y_offsetcount)
                cn.pe_vb_katauke_bend_90_down_st_ba(msp, offset_x=x_count, offset_y=y_count)
                first_x = x_count
                x_count += 8000000
                y_count += - 10000000
                y_offsetcount += 1
                today_count += 1
                second_x = x_count
                flat_sum += (item.get('水平距離', 0) or 0) + (item.get('水平距離2', 0) or 0)
                haikann_sum += (item.get('配管長', 0) or 0) + (item.get('配管長2', 0) or 0)
                msp.add_line((first_x, y_count - y_count - 45000000), (first_x, y_count - y_count - 61000000))
                msp.add_line((second_x, y_count - y_count - 45000000), (second_x, y_count - y_count - 61000000))
                msp.add_line((first_x, y_count - y_count - 61000000), (second_x, y_count - y_count - 61000000))
                cn.flat_line1(msp, item, offset_x=x_count-4000000, offset_y=y_count - y_count - 61000000)
        elif item['方向'] == 'AB':
            if today_count == 0:
                cn.pe_vb_katauke_bend_90_down_st_ab(msp, offset_x=x_count, offset_y=y_count)
                x_count += 10000000
                y_count += -8000000
                today_count += 1
                today_xbase += 10000000
            else:
                cn.choose_zairyou_no(msp, item, offset_x=x_count + 5000000, offset_y=y_count + 15000000)
                cn.name_sokutenn(msp, item, route_name,offset_x=x_count - 90000000, offset_y=y_count + 8000000 + 13500000 * y_offsetcount)
                cn.pe_vb_katauke_bend_90_down_st_ab(msp, offset_x=x_count, offset_y=y_count)
                first_x = x_count
                x_count += 10000000
                y_count += -8000000
                y_offsetcount += 1
                today_count += 1
                second_x = x_count
                flat_sum += (item.get('水平距離', 0) or 0) + (item.get('水平距離2', 0) or 0)
                haikann_sum += (item.get('配管長', 0) or 0) + (item.get('配管長2', 0) or 0)
                msp.add_line((first_x, y_count - y_count - 45000000), (first_x, y_count - y_count - 61000000))
                msp.add_line((second_x, y_count - y_count - 45000000), (second_x, y_count - y_count - 61000000))
                msp.add_line((first_x, y_count - y_count - 61000000), (second_x, y_count - y_count - 61000000))
                cn.flat_line1(msp, item, offset_x=x_count-5000000, offset_y=y_count - y_count - 61000000)
    elif (
        'EF' in name
        and 'ベンド' in name
        and '90°' in name
        and '(V)' in name
        and '上がり' in name
    ):
        if item['方向'] == 'AA':
            if today_count == 0:
                cn.pe_vb_ryouuke_bend_90_up_st_aa(msp, offset_x=x_count, offset_y=y_count)
                x_count += 10000000
                y_count += 10000000
                today_count += 1
                today_xbase += 10000000
            else:
                cn.choose_zairyou_no(msp, item, offset_x=x_count, offset_y=y_count + 15000000)
                cn.name_sokutenn(msp, item, route_name,offset_x=x_count - 90000000, offset_y=y_count + 8000000 + 13500000 * y_offsetcount)
                cn.pe_vb_ryouuke_bend_90_up_st_aa(msp, offset_x=x_count, offset_y=y_count)
                first_x = x_count
                x_count += 10000000
                y_count += 10000000
                y_offsetcount += 1
                today_count += 1
                second_x = x_count
                flat_sum += (item.get('水平距離', 0) or 0) + (item.get('水平距離2', 0) or 0)
                haikann_sum += (item.get('配管長', 0) or 0) + (item.get('配管長2', 0) or 0)
                msp.add_line((first_x, y_count - y_count - 45000000), (first_x, y_count - y_count - 61000000))
                msp.add_line((second_x, y_count - y_count - 45000000), (second_x, y_count - y_count - 61000000))
                msp.add_line((first_x, y_count - y_count - 61000000), (second_x, y_count - y_count - 61000000))
                cn.flat_line1(msp, item, offset_x=x_count-5000000, offset_y=y_count - y_count - 61000000)
        elif item['方向'] == 'AB':
            if today_count == 0:
                cn.pe_vb_katauke_bend_90_up_st_ab(msp, offset_x=x_count, offset_y=y_count)
                x_count += 10000000
                y_count += 8000000
                today_count += 1
                today_xbase += 10000000
            else:
                cn.choose_zairyou_no(msp, item, offset_x=x_count, offset_y=y_count + 15000000)
                cn.name_sokutenn(msp, item, route_name,offset_x=x_count - 90000000, offset_y=y_count + 8000000 + 13500000 * y_offsetcount)
                cn.pe_vb_katauke_bend_90_up_st_ab(msp, offset_x=x_count, offset_y=y_count)
                first_x = x_count
                x_count += 10000000
                y_count += 8000000
                y_offsetcount += 1
                today_count += 1
                second_x = x_count
                flat_sum += (item.get('水平距離', 0) or 0) + (item.get('水平距離2', 0) or 0)
                haikann_sum += (item.get('配管長', 0) or 0) + (item.get('配管長2', 0) or 0)
                msp.add_line((first_x, y_count - y_count - 45000000), (first_x, y_count - y_count - 61000000))
                msp.add_line((second_x, y_count - y_count - 45000000), (second_x, y_count - y_count - 61000000))
                msp.add_line((first_x, y_count - y_count - 61000000), (second_x, y_count - y_count - 61000000))
                cn.flat_line1(msp, item, offset_x=x_count-5000000, offset_y=y_count - y_count - 61000000)                
        elif item['方向'] == 'BA':
            if today_count == 0:
                cn.pe_vb_katauke_bend_90_up_st_ba(msp, offset_x=x_count, offset_y=y_count)
                x_count += 8000000
                y_count += 10000000
                today_count += 1
                today_xbase += 8000000
            else:
                cn.choose_zairyou_no(msp, item, offset_x=x_count - 3000000, offset_y=y_count + 15000000)
                cn.name_sokutenn(msp, item, route_name,offset_x=x_count - 90000000, offset_y=y_count + 8000000 + 13500000 * y_offsetcount)
                cn.pe_vb_katauke_bend_90_up_st_ba(msp, offset_x=x_count, offset_y=y_count)
                first_x = x_count
                x_count += 8000000
                y_count += 10000000
                y_offsetcount += 1
                today_count += 1
                second_x = x_count
                flat_sum += (item.get('水平距離', 0) or 0) + (item.get('水平距離2', 0) or 0)
                haikann_sum += (item.get('配管長', 0) or 0) + (item.get('配管長2', 0) or 0)
                msp.add_line((first_x, y_count - y_count - 45000000), (first_x, y_count - y_count - 61000000))
                msp.add_line((second_x, y_count - y_count - 45000000), (second_x, y_count - y_count - 61000000))
                msp.add_line((first_x, y_count - y_count - 61000000), (second_x, y_count - y_count - 61000000))
                cn.flat_line1(msp, item, offset_x=x_count - 4000000, offset_y=y_count - y_count - 61000000) 
                
    elif (
        'EF' in name
        and 'ベンド' in name
        and '45°' in name
        and '(V)' in name
        and '下がり' in name
    ):
        if item['方向'] == 'AA':
            if today_count == 0:
                cn.pe_vb_ryouuke_bend_45_down_st_aa(msp, offset_x=x_count, offset_y=y_count)
                x_count += 17000000
                y_count += -7000000
                today_count += 1
                today_xbase += 17000000
            else:
                cn.choose_zairyou_no(msp, item, offset_x=x_count + 10000000, offset_y=y_count + 15000000)
                cn.name_sokutenn(msp, item, route_name,offset_x=x_count - 90000000, offset_y=y_count - y_count + 13500000 * y_offsetcount)
                cn.pe_vb_ryouuke_bend_45_down_st_aa(msp, offset_x=x_count, offset_y=y_count)
                first_x = x_count
                x_count += 17000000
                y_count += -7000000
                y_offsetcount += 1
                today_count += 1
                second_x = x_count
                flat_sum += (item.get('水平距離', 0) or 0) + (item.get('水平距離2', 0) or 0)
                haikann_sum += (item.get('配管長', 0) or 0) + (item.get('配管長2', 0) or 0)
                msp.add_line((first_x, y_count - y_count - 45000000), (first_x, y_count - y_count - 61000000))
                msp.add_line((first_x + 10000000, y_count - y_count - 45000000), (first_x + 10000000, y_count - y_count - 61000000))
                msp.add_line((second_x, y_count - y_count - 45000000), (second_x, y_count - y_count - 61000000))
                msp.add_line((first_x, y_count - y_count - 61000000), (second_x, y_count - y_count - 61000000))
                cn.flat_line1(msp, item, offset_x=x_count-12000000, offset_y=y_count - y_count - 61000000)
                cn.flat_line2(msp, item, offset_x=x_count-3500000, offset_y=y_count - y_count - 61000000)
                if item['水平距離'] == item['配管長']:
                    pass
                else:
                    cn.haikann_line1(msp, item, offset_x=x_count-19500000, offset_y=y_count - y_count - 54500000)                   
                if item['水平距離2'] == item['配管長2']:
                    pass
                else:
                    cn.haikann_line2(msp, item, offset_x=x_count-11000000, offset_y=y_count - y_count - 54500000)
        elif item['方向'] == 'AB':
            if today_count == 0:
                cn.pe_vb_katauke_bend_45_down_st_ab(msp, offset_x=x_count, offset_y=y_count)
                x_count += 15660000
                y_count += -5660000
                today_count += 1
                today_xbase += 15660000
            else:
                cn.choose_zairyou_no(msp, item, offset_x=x_count + 8660000, offset_y=y_count + 15000000)
                cn.name_sokutenn(msp, item, route_name,offset_x=x_count - 90000000, offset_y=y_count - y_count + 13500000 * y_offsetcount)
                cn.pe_vb_katauke_bend_45_down_st_ab(msp, offset_x=x_count, offset_y=y_count)
                first_x = x_count
                x_count += 15660000
                y_count += -5660000
                y_offsetcount += 1
                today_count += 1
                second_x = x_count
                flat_sum += (item.get('水平距離', 0) or 0) + (item.get('水平距離2', 0) or 0)
                haikann_sum += (item.get('配管長', 0) or 0) + (item.get('配管長2', 0) or 0)
                msp.add_line((first_x, y_count - y_count - 45000000), (first_x, y_count - y_count - 61000000))
                msp.add_line((first_x + 10000000, y_count - y_count - 45000000), (first_x + 10000000, y_count - y_count - 61000000))
                msp.add_line((second_x, y_count - y_count - 45000000), (second_x, y_count - y_count - 61000000))
                msp.add_line((first_x, y_count - y_count - 61000000), (second_x, y_count - y_count - 61000000))
                cn.flat_line1(msp, item, offset_x=x_count-12000000, offset_y=y_count - y_count - 61000000)
                cn.flat_line2(msp, item, offset_x=x_count-3500000, offset_y=y_count - y_count - 61000000)
                if item['水平距離'] == item['配管長']:
                    pass
                else:
                    cn.haikann_line1(msp, item, offset_x=x_count-19500000, offset_y=y_count - y_count - 54500000)                   
                if item['水平距離2'] == item['配管長2']:
                    pass
                else:
                    cn.haikann_line2(msp, item, offset_x=x_count-11000000, offset_y=y_count - y_count - 54500000)
        elif item['方向'] == 'BA':
            if today_count == 0:
                cn.pe_vb_katauke_bend_45_down_st_ba(msp, offset_x=x_count, offset_y=y_count)
                x_count += 15080000
                y_count += - 7080000
                today_count += 1
                today_xbase += 15080000
            else:
                cn.choose_zairyou_no(msp, item, offset_x=x_count + 8660000, offset_y=y_count + 15000000)
                cn.name_sokutenn(msp, item, route_name,offset_x=x_count - 90000000, offset_y=y_count - y_count + 13500000 * y_offsetcount)
                cn.pe_vb_katauke_bend_45_down_st_ba(msp, offset_x=x_count, offset_y=y_count)
                first_x = x_count
                x_count += 15080000
                y_count += -7080000
                y_offsetcount += 1
                today_count += 1
                second_x = x_count
                flat_sum += (item.get('水平距離', 0) or 0) + (item.get('水平距離2', 0) or 0)
                haikann_sum += (item.get('配管長', 0) or 0) + (item.get('配管長2', 0) or 0)
                msp.add_line((first_x, y_count - y_count - 45000000), (first_x, y_count - y_count - 61000000))
                msp.add_line((first_x + 10000000, y_count - y_count - 45000000), (first_x + 10000000, y_count - y_count - 61000000))
                msp.add_line((second_x, y_count - y_count - 45000000), (second_x, y_count - y_count - 61000000))
                msp.add_line((first_x, y_count - y_count - 61000000), (second_x, y_count - y_count - 61000000))
                cn.flat_line1(msp, item, offset_x=x_count-12000000, offset_y=y_count - y_count - 61000000)
                cn.flat_line2(msp, item, offset_x=x_count-3500000, offset_y=y_count - y_count - 61000000)
                if item['水平距離'] == item['配管長']:
                    pass
                else:
                    cn.haikann_line1(msp, item, offset_x=x_count-19500000, offset_y=y_count - y_count - 54500000)                   
                if item['水平距離2'] == item['配管長2']:
                    pass
                else:
                    cn.haikann_line2(msp, item, offset_x=x_count-11000000, offset_y=y_count - y_count - 54500000)                    
    elif (
        'EF' in name
        and 'ベンド' in name
        and '45°' in name
        and '(V)' in name
        and '上がり' in name
    ):
        if item['方向'] == 'AA':
            if today_count == 0:
                cn.pe_vb_ryouuke_bend_45_up_st_aa(msp, offset_x=x_count, offset_y=y_count)
                x_count += 17000000
                y_count += 7000000
                today_count += 1
                today_xbase += 17000000
            else:
                cn.choose_zairyou_no(msp, item, offset_x=x_count + 8500000, offset_y=y_count + 15000000)
                cn.name_sokutenn(msp, item, route_name,offset_x=x_count - 90000000, offset_y=y_count - y_count + 13500000 * y_offsetcount)
                cn.pe_vb_ryouuke_bend_45_up_st_aa(msp, offset_x=x_count, offset_y=y_count)
                first_x = x_count
                x_count += 17000000
                y_count += 7000000
                y_offsetcount += 1
                today_count += 1
                second_x = x_count
                flat_sum += (item.get('水平距離', 0) or 0) + (item.get('水平距離2', 0) or 0)
                haikann_sum += (item.get('配管長', 0) or 0) + (item.get('配管長2', 0) or 0)
                msp.add_line((first_x, y_count - y_count - 45000000), (first_x, y_count - y_count - 61000000))
                msp.add_line((first_x + 10000000, y_count - y_count - 45000000), (first_x + 10000000, y_count - y_count - 61000000))
                msp.add_line((second_x, y_count - y_count - 45000000), (second_x, y_count - y_count - 61000000))
                msp.add_line((first_x, y_count - y_count - 61000000), (second_x, y_count - y_count - 61000000))
                cn.flat_line1(msp, item, offset_x=x_count-12000000, offset_y=y_count - y_count - 61000000)
                cn.flat_line2(msp, item, offset_x=x_count-3500000, offset_y=y_count - y_count - 61000000)
                if item['水平距離'] == item['配管長']:
                    pass
                else:
                    cn.haikann_line1(msp, item, offset_x=x_count-19500000, offset_y=y_count - y_count - 54500000)                   
                if item['水平距離2'] == item['配管長2']:
                    pass
                else:
                    cn.haikann_line2(msp, item, offset_x=x_count-11000000, offset_y=y_count - y_count - 54500000)
        elif item['方向'] == 'AB':
            if today_count == 0:
                cn.pe_vb_katauke_bend_45_up_st_ab(msp, offset_x=x_count, offset_y=y_count)
                x_count += 15660000
                y_count += 5660000
                today_count += 1
                today_xbase += 15660000
            else:
                cn.choose_zairyou_no(msp, item, offset_x=x_count + 2000000, offset_y=y_count + 15000000)
                cn.name_sokutenn(msp, item, route_name,offset_x=x_count - 90000000, offset_y=y_count - y_count + 13500000 * y_offsetcount)
                cn.pe_vb_katauke_bend_45_up_st_ab(msp, offset_x=x_count, offset_y=y_count)
                first_x = x_count
                x_count += 15660000
                y_count += 5660000
                y_offsetcount += 1
                today_count += 1
                second_x = x_count
                flat_sum += (item.get('水平距離', 0) or 0) + (item.get('水平距離2', 0) or 0)
                haikann_sum += (item.get('配管長', 0) or 0) + (item.get('配管長2', 0) or 0)
                msp.add_line((first_x, y_count - y_count - 45000000), (first_x, y_count - y_count - 61000000))
                msp.add_line((first_x + 10000000, y_count - y_count - 45000000), (first_x + 10000000, y_count - y_count - 61000000))
                msp.add_line((second_x, y_count - y_count - 45000000), (second_x, y_count - y_count - 61000000))
                msp.add_line((first_x, y_count - y_count - 61000000), (second_x, y_count - y_count - 61000000))
                cn.flat_line1(msp, item, offset_x=x_count-12000000, offset_y=y_count - y_count - 61000000)
                cn.flat_line2(msp, item, offset_x=x_count-3500000, offset_y=y_count - y_count - 61000000)
                if item['水平距離'] == item['配管長']:
                    pass
                else:
                    cn.haikann_line1(msp, item, offset_x=x_count-19500000, offset_y=y_count - y_count - 54500000)                   
                if item['水平距離2'] == item['配管長2']:
                    pass
                else:
                    cn.haikann_line2(msp, item, offset_x=x_count-11000000, offset_y=y_count - y_count - 54500000)
        elif item['方向'] == 'BA':
            if today_count == 0:
                cn.pe_vb_katauke_bend_45_up_st_ba(msp, offset_x=x_count, offset_y=y_count)
                x_count += 15080000
                y_count += 7080000
                today_count += 1
                today_xbase += 15080000
            else:
                cn.choose_zairyou_no(msp, item, offset_x=x_count + 2000000, offset_y=y_count + 15000000)
                cn.name_sokutenn(msp, item, route_name,offset_x=x_count - 90000000, offset_y=y_count - y_count + 13500000 * y_offsetcount)
                cn.pe_vb_katauke_bend_45_up_st_ba(msp, offset_x=x_count, offset_y=y_count)
                first_x = x_count
                x_count += 15080000
                y_count += 7080000
                y_offsetcount += 1
                today_count += 1
                second_x = x_count
                flat_sum += (item.get('水平距離', 0) or 0) + (item.get('水平距離2', 0) or 0)
                haikann_sum += (item.get('配管長', 0) or 0) + (item.get('配管長2', 0) or 0)
                msp.add_line((first_x, y_count - y_count - 45000000), (first_x, y_count - y_count - 61000000))
                msp.add_line((first_x + 10000000, y_count - y_count - 45000000), (first_x + 10000000, y_count - y_count - 61000000))
                msp.add_line((second_x, y_count - y_count - 45000000), (second_x, y_count - y_count - 61000000))
                msp.add_line((first_x, y_count - y_count - 61000000), (second_x, y_count - y_count - 61000000))
                cn.flat_line1(msp, item, offset_x=x_count-12000000, offset_y=y_count - y_count - 61000000)
                cn.flat_line2(msp, item, offset_x=x_count-3500000, offset_y=y_count - y_count - 61000000)
                if item['水平距離'] == item['配管長']:
                    pass
                else:
                    cn.haikann_line1(msp, item, offset_x=x_count-19500000, offset_y=y_count - y_count - 54500000)                   
                if item['水平距離2'] == item['配管長2']:
                    pass
                else:
                    cn.haikann_line2(msp, item, offset_x=x_count-11000000, offset_y=y_count - y_count - 54500000)                   
    elif (
        'EF' in name
        and 'ベンド' in name
        and '22' in name
        and '(V)' in name
        and '下がり' in name
    ):
        if item['方向'] == 'AA':
            if today_count == 0:
                cn.pe_vb_ryouuke_bend_22_down_st_aa(msp, offset_x=x_count, offset_y=y_count)
                x_count += 19240000
                y_count += - 3820000
                today_count += 1
                today_xbase += 19240000
            else:
                cn.choose_zairyou_no(msp, item, offset_x=x_count + 10000000, offset_y=y_count + 15000000)
                cn.name_sokutenn(msp, item, route_name,offset_x=x_count - 90000000, offset_y=y_count - y_count + 13500000 * y_offsetcount)
                cn.pe_vb_ryouuke_bend_22_down_st_aa(msp, offset_x=x_count, offset_y=y_count)
                first_x = x_count
                x_count += 19240000
                y_count += - 3820000
                y_offsetcount += 1
                today_count += 1
                second_x = x_count
                flat_sum += (item.get('水平距離', 0) or 0) + (item.get('水平距離2', 0) or 0)
                haikann_sum += (item.get('配管長', 0) or 0) + (item.get('配管長2', 0) or 0)
                msp.add_line((first_x, y_count - y_count - 45000000), (first_x, y_count - y_count - 61000000))
                msp.add_line((first_x + 10000000, y_count - y_count - 45000000), (first_x + 10000000, y_count - y_count - 61000000))
                msp.add_line((second_x, y_count - y_count - 45000000), (second_x, y_count - y_count - 61000000))
                msp.add_line((first_x, y_count - y_count - 61000000), (second_x, y_count - y_count - 61000000))
                cn.flat_line1(msp, item, offset_x=x_count-14200000, offset_y=y_count - y_count - 61000000)
                cn.flat_line2(msp, item, offset_x=x_count-4700000, offset_y=y_count - y_count - 61000000)
                if item['水平距離'] == item['配管長']:
                    pass
                else:
                    cn.haikann_line1(msp, item, offset_x=x_count-19500000, offset_y=y_count - y_count - 54500000)                   
                if item['水平距離2'] == item['配管長2']:
                    pass
                else:
                    cn.haikann_line2(msp, item, offset_x=x_count-12100000, offset_y=y_count - y_count - 54500000)
        elif item['方向'] == 'AB':
            if today_count == 0:
                cn.pe_vb_katauke_bend_22_down_st_ab(msp, offset_x=x_count, offset_y=y_count)
                x_count += 17380000
                y_count += - 3060000
                today_count += 1
                today_xbase += 17380000
            else:
                cn.choose_zairyou_no(msp, item, offset_x=x_count + 10000000, offset_y=y_count + 15000000)
                cn.name_sokutenn(msp, item, route_name,offset_x=x_count - 90000000, offset_y=y_count - y_count + 13500000 * y_offsetcount)
                cn.pe_vb_katauke_bend_22_down_st_ab(msp, offset_x=x_count, offset_y=y_count)
                first_x = x_count
                x_count += 17380000
                y_count += - 3060000
                y_offsetcount += 1
                today_count += 1
                second_x = x_count
                flat_sum += (item.get('水平距離', 0) or 0) + (item.get('水平距離2', 0) or 0)
                haikann_sum += (item.get('配管長', 0) or 0) + (item.get('配管長2', 0) or 0)
                msp.add_line((first_x, y_count - y_count - 45000000), (first_x, y_count - y_count - 61000000))
                msp.add_line((first_x + 10000000, y_count - y_count - 45000000), (first_x + 10000000, y_count - y_count - 61000000))
                msp.add_line((second_x, y_count - y_count - 45000000), (second_x, y_count - y_count - 61000000))
                msp.add_line((first_x, y_count - y_count - 61000000), (second_x, y_count - y_count - 61000000))
                cn.flat_line1(msp, item, offset_x=x_count-12350000, offset_y=y_count - y_count - 61000000)
                cn.flat_line2(msp, item, offset_x=x_count-3700000, offset_y=y_count - y_count - 61000000)
                if item['水平距離'] == item['配管長']:
                    pass
                else:
                    cn.haikann_line1(msp, item, offset_x=x_count-19500000, offset_y=y_count - y_count - 54500000)                   
                if item['水平距離2'] == item['配管長2']:
                    pass
                else:
                    cn.haikann_line2(msp, item, offset_x=x_count-11100000, offset_y=y_count - y_count - 54500000)
        elif item['方向'] == 'BA':
            if today_count == 0:
                cn.pe_vb_katauke_bend_22_down_st_ba(msp, offset_x=x_count, offset_y=y_count)
                x_count += 17240000
                y_count += - 3820000
                today_count += 1
                today_xbase += 17240000
            else:
                cn.choose_zairyou_no(msp, item, offset_x=x_count + 8000000, offset_y=y_count + 15000000)
                cn.name_sokutenn(msp, item, route_name,offset_x=x_count - 90000000, offset_y=y_count - y_count + 13500000 * y_offsetcount)
                cn.pe_vb_katauke_bend_22_down_st_ba(msp, offset_x=x_count, offset_y=y_count)
                first_x = x_count
                x_count += 17240000
                y_count += - 3820000
                y_offsetcount += 1
                today_count += 1
                second_x = x_count
                flat_sum += (item.get('水平距離', 0) or 0) + (item.get('水平距離2', 0) or 0)
                haikann_sum += (item.get('配管長', 0) or 0) + (item.get('配管長2', 0) or 0)
                msp.add_line((first_x, y_count - y_count - 45000000), (first_x, y_count - y_count - 61000000))
                msp.add_line((first_x + 8000000, y_count - y_count - 45000000), (first_x + 8000000, y_count - y_count - 61000000))
                msp.add_line((second_x, y_count - y_count - 45000000), (second_x, y_count - y_count - 61000000))
                msp.add_line((first_x, y_count - y_count - 61000000), (second_x, y_count - y_count - 61000000))
                cn.flat_line1(msp, item, offset_x=x_count-12350000, offset_y=y_count - y_count - 61000000)
                cn.flat_line2(msp, item, offset_x=x_count-3700000, offset_y=y_count - y_count - 61000000)
                if item['水平距離'] == item['配管長']:
                    pass
                else:
                    cn.haikann_line1(msp, item, offset_x=x_count-19500000, offset_y=y_count - y_count - 54500000)                   
                if item['水平距離2'] == item['配管長2']:
                    pass
                else:
                    cn.haikann_line2(msp, item, offset_x=x_count-11100000, offset_y=y_count - y_count - 54500000)
    elif (
        'EF' in name
        and 'ベンド' in name
        and '22' in name
        and '(V)' in name
        and '上がり' in name
    ):
        if item['方向'] == 'AA':
            if today_count == 0:
                cn.pe_vb_ryouuke_bend_22_up_st_aa(msp, offset_x=x_count, offset_y=y_count)
                x_count += 19240000
                y_count += 3820000
                today_count += 1
                today_xbase += 19240000
            else:
                cn.choose_zairyou_no(msp, item, offset_x=x_count + 10000000, offset_y=y_count + 15000000)
                cn.name_sokutenn(msp, item, route_name,offset_x=x_count - 90000000, offset_y=y_count - y_count + 13500000 * y_offsetcount)
                cn.pe_vb_ryouuke_bend_22_up_st_aa(msp, offset_x=x_count, offset_y=y_count)
                first_x = x_count
                x_count += 19240000
                y_count += 3820000
                y_offsetcount += 1
                today_count += 1
                second_x = x_count
                flat_sum += (item.get('水平距離', 0) or 0) + (item.get('水平距離2', 0) or 0)
                haikann_sum += (item.get('配管長', 0) or 0) + (item.get('配管長2', 0) or 0)
                msp.add_line((first_x, y_count - y_count - 45000000), (first_x, y_count - y_count - 61000000))
                msp.add_line((first_x + 10000000, y_count - y_count - 45000000), (first_x + 10000000, y_count - y_count - 61000000))
                msp.add_line((second_x, y_count - y_count - 45000000), (second_x, y_count - y_count - 61000000))
                msp.add_line((first_x, y_count - y_count - 61000000), (second_x, y_count - y_count - 61000000))
                cn.flat_line1(msp, item, offset_x=x_count-14200000, offset_y=y_count - y_count - 61000000)
                cn.flat_line2(msp, item, offset_x=x_count-4700000, offset_y=y_count - y_count - 61000000)
                if item['水平距離'] == item['配管長']:
                    pass
                else:
                    cn.haikann_line1(msp, item, offset_x=x_count-19500000, offset_y=y_count - y_count - 54500000)                   
                if item['水平距離2'] == item['配管長2']:
                    pass
                else:
                    cn.haikann_line2(msp, item, offset_x=x_count-12100000, offset_y=y_count - y_count - 54500000)
        elif item['方向'] == 'AB':
            if today_count == 0:
                cn.pe_vb_katauke_bend_22_up_st_ab(msp, offset_x=x_count, offset_y=y_count)
                x_count += 17380000
                y_count += 3060000
                today_count += 1
                today_xbase += 17380000
            else:
                cn.choose_zairyou_no(msp, item, offset_x=x_count + 10000000, offset_y=y_count + 15000000)
                cn.name_sokutenn(msp, item, route_name,offset_x=x_count - 90000000, offset_y=y_count - y_count + 13500000 * y_offsetcount)
                cn.pe_vb_katauke_bend_22_up_st_ab(msp, offset_x=x_count, offset_y=y_count)
                first_x = x_count
                x_count += 17380000
                y_count += 3060000
                y_offsetcount += 1
                today_count += 1
                second_x = x_count
                flat_sum += (item.get('水平距離', 0) or 0) + (item.get('水平距離2', 0) or 0)
                haikann_sum += (item.get('配管長', 0) or 0) + (item.get('配管長2', 0) or 0)
                msp.add_line((first_x, y_count - y_count - 45000000), (first_x, y_count - y_count - 61000000))
                msp.add_line((first_x + 10000000, y_count - y_count - 45000000), (first_x + 10000000, y_count - y_count - 61000000))
                msp.add_line((second_x, y_count - y_count - 45000000), (second_x, y_count - y_count - 61000000))
                msp.add_line((first_x, y_count - y_count - 61000000), (second_x, y_count - y_count - 61000000))
                cn.flat_line1(msp, item, offset_x=x_count-12400000, offset_y=y_count - y_count - 61000000)
                cn.flat_line2(msp, item, offset_x=x_count-3700000, offset_y=y_count - y_count - 61000000)
                if item['水平距離'] == item['配管長']:
                    pass
                else:
                    cn.haikann_line1(msp, item, offset_x=x_count-19500000, offset_y=y_count - y_count - 54500000)                   
                if item['水平距離2'] == item['配管長2']:
                    pass
                else:
                    cn.haikann_line2(msp, item, offset_x=x_count-11100000, offset_y=y_count - y_count - 54500000)
        elif item['方向'] == 'BA':
            if today_count == 0:
                cn.pe_vb_katauke_bend_22_up_st_ba(msp, offset_x=x_count, offset_y=y_count)
                x_count += 17240000
                y_count += 3820000
                today_count += 1
                today_xbase += 17380000
            else:
                cn.choose_zairyou_no(msp, item, offset_x=x_count + 8000000, offset_y=y_count + 15000000)
                cn.name_sokutenn(msp, item, route_name,offset_x=x_count - 90000000, offset_y=y_count - y_count + 13500000 * y_offsetcount)
                cn.pe_vb_katauke_bend_22_up_st_ba(msp, offset_x=x_count, offset_y=y_count)
                first_x = x_count
                x_count += 17240000
                y_count += 3820000
                y_offsetcount += 1
                today_count += 1
                second_x = x_count
                flat_sum += (item.get('水平距離', 0) or 0) + (item.get('水平距離2', 0) or 0)
                haikann_sum += (item.get('配管長', 0) or 0) + (item.get('配管長2', 0) or 0)
                msp.add_line((first_x, y_count - y_count - 45000000), (first_x, y_count - y_count - 61000000))
                msp.add_line((first_x + 10000000, y_count - y_count - 45000000), (first_x + 10000000, y_count - y_count - 61000000))
                msp.add_line((second_x, y_count - y_count - 45000000), (second_x, y_count - y_count - 61000000))
                msp.add_line((first_x, y_count - y_count - 61000000), (second_x, y_count - y_count - 61000000))
                cn.flat_line1(msp, item, offset_x=x_count-12400000, offset_y=y_count - y_count - 61000000)
                cn.flat_line2(msp, item, offset_x=x_count-3700000, offset_y=y_count - y_count - 61000000)
                if item['水平距離'] == item['配管長']:
                    pass
                else:
                    cn.haikann_line1(msp, item, offset_x=x_count-19500000, offset_y=y_count - y_count - 54500000)                   
                if item['水平距離2'] == item['配管長2']:
                    pass
                else:
                    cn.haikann_line2(msp, item, offset_x=x_count-11100000, offset_y=y_count - y_count - 54500000)
    elif (
        'EF' in name
        and 'ベンド' in name
        and '11' in name
        and '(V)' in name
        and '下がり' in name
    ):
        if item['方向'] == 'AA':
            if today_count == 0:
                cn.pe_vb_ryouuke_bend_11_down_st_aa(msp, offset_x=x_count, offset_y=y_count)
                x_count += 19800000
                y_count += - 1960000
                today_count += 1
                today_xbase += 19800000
            else:
                cn.choose_zairyou_no(msp, item, offset_x=x_count + 9900000, offset_y=y_count + 15000000)
                cn.name_sokutenn(msp, item, route_name,offset_x=x_count - 90000000, offset_y=y_count - y_count + 13500000 * y_offsetcount)
                cn.pe_vb_ryouuke_bend_11_down_st_aa(msp, offset_x=x_count, offset_y=y_count)
                first_x = x_count
                x_count += 19800000
                y_count += - 1960000
                y_offsetcount += 1
                today_count += 1
                second_x = x_count
                flat_sum += (item.get('水平距離', 0) or 0) + (item.get('水平距離2', 0) or 0)
                haikann_sum += (item.get('配管長', 0) or 0) + (item.get('配管長2', 0) or 0)
                msp.add_line((first_x, y_count - y_count - 45000000), (first_x, y_count - y_count - 61000000))
                msp.add_line((first_x + 10000000, y_count - y_count - 45000000), (first_x + 10000000, y_count - y_count - 61000000))
                msp.add_line((second_x, y_count - y_count - 45000000), (second_x, y_count - y_count - 61000000))
                msp.add_line((first_x, y_count - y_count - 61000000), (second_x, y_count - y_count - 61000000))
                cn.flat_line1(msp, item, offset_x=x_count-14200000, offset_y=y_count - y_count - 61000000)
                cn.flat_line2(msp, item, offset_x=x_count-4700000, offset_y=y_count - y_count - 61000000)
                if item['水平距離'] == item['配管長']:
                    pass
                else:
                    cn.haikann_line1(msp, item, offset_x=x_count-19500000, offset_y=y_count - y_count - 54500000)                   
                if item['水平距離2'] == item['配管長2']:
                    pass
                else:
                    cn.haikann_line2(msp, item, offset_x=x_count-12100000, offset_y=y_count - y_count - 54500000)
        elif item['方向'] == 'AB':
            if today_count == 0:
                cn.pe_vb_katauke_bend_11_down_st_ab(msp, offset_x=x_count, offset_y=y_count)
                x_count += 17840000
                y_count += - 1570000
                today_count += 1
                today_xbase += 17840000
            else:
                cn.choose_zairyou_no(msp, item, offset_x=x_count + 10000000, offset_y=y_count + 15000000)
                cn.name_sokutenn(msp, item, route_name,offset_x=x_count - 90000000, offset_y=y_count - y_count + 13500000 * y_offsetcount)
                cn.pe_vb_katauke_bend_11_down_st_ab(msp, offset_x=x_count, offset_y=y_count)
                first_x = x_count
                x_count += 17840000
                y_count += - 1570000
                y_offsetcount += 1
                today_count += 1
                second_x = x_count
                flat_sum += (item.get('水平距離', 0) or 0) + (item.get('水平距離2', 0) or 0)
                haikann_sum += (item.get('配管長', 0) or 0) + (item.get('配管長2', 0) or 0)
                msp.add_line((first_x, y_count - y_count - 45000000), (first_x, y_count - y_count - 61000000))
                msp.add_line((first_x + 10000000, y_count - y_count - 45000000), (first_x + 10000000, y_count - y_count - 61000000))
                msp.add_line((second_x, y_count - y_count - 45000000), (second_x, y_count - y_count - 61000000))
                msp.add_line((first_x, y_count - y_count - 61000000), (second_x, y_count - y_count - 61000000))
                cn.flat_line1(msp, item, offset_x=x_count-12350000, offset_y=y_count - y_count - 61000000)
                cn.flat_line2(msp, item, offset_x=x_count-3700000, offset_y=y_count - y_count - 61000000)
                if item['水平距離'] == item['配管長']:
                    pass
                else:
                    cn.haikann_line1(msp, item, offset_x=x_count-19500000, offset_y=y_count - y_count - 54500000)                   
                if item['水平距離2'] == item['配管長2']:
                    pass
                else:
                    cn.haikann_line2(msp, item, offset_x=x_count-11100000, offset_y=y_count - y_count - 54500000)
        elif item['方向'] == 'BA':
            if today_count == 0:
                cn.pe_vb_katauke_bend_11_down_st_ba(msp, offset_x=x_count, offset_y=y_count)
                x_count += 17800000
                y_count += - 1960000
                today_count += 1
                today_xbase += 17800000
            else:
                cn.choose_zairyou_no(msp, item, offset_x=x_count + 8000000, offset_y=y_count + 15000000)
                cn.name_sokutenn(msp, item, route_name,offset_x=x_count - 90000000, offset_y=y_count - y_count + 13500000 * y_offsetcount)
                cn.pe_vb_katauke_bend_11_down_st_ba(msp, offset_x=x_count, offset_y=y_count)
                first_x = x_count
                x_count += 17800000
                y_count += - 1960000
                y_offsetcount += 1
                today_count += 1
                second_x = x_count
                flat_sum += (item.get('水平距離', 0) or 0) + (item.get('水平距離2', 0) or 0)
                haikann_sum += (item.get('配管長', 0) or 0) + (item.get('配管長2', 0) or 0)
                msp.add_line((first_x, y_count - y_count - 45000000), (first_x, y_count - y_count - 61000000))
                msp.add_line((first_x + 8000000, y_count - y_count - 45000000), (first_x + 8000000, y_count - y_count - 61000000))
                msp.add_line((second_x, y_count - y_count - 45000000), (second_x, y_count - y_count - 61000000))
                msp.add_line((first_x, y_count - y_count - 61000000), (second_x, y_count - y_count - 61000000))
                cn.flat_line1(msp, item, offset_x=x_count-12350000, offset_y=y_count - y_count - 61000000)
                cn.flat_line2(msp, item, offset_x=x_count-3700000, offset_y=y_count - y_count - 61000000)
                if item['水平距離'] == item['配管長']:
                    pass
                else:
                    cn.haikann_line1(msp, item, offset_x=x_count-19500000, offset_y=y_count - y_count - 54500000)                   
                if item['水平距離2'] == item['配管長2']:
                    pass
                else:
                    cn.haikann_line2(msp, item, offset_x=x_count-11100000, offset_y=y_count - y_count - 54500000)
    elif (
        'EF' in name
        and 'ベンド' in name
        and '11' in name
        and '(V)' in name
        and '上がり' in name
    ):
        if item['方向'] == 'AA':
            if today_count == 0:
                cn.pe_vb_ryouuke_bend_11_up_st_aa(msp, offset_x=x_count, offset_y=y_count)
                x_count += 19800000
                y_count += 1960000
                today_count += 1
                today_xbase += 19800000
            else:
                cn.choose_zairyou_no(msp, item, offset_x=x_count + 10000000, offset_y=y_count + 15000000)
                cn.name_sokutenn(msp, item, route_name,offset_x=x_count - 90000000, offset_y=y_count - y_count + 13500000 * y_offsetcount)
                cn.pe_vb_ryouuke_bend_11_up_st_aa(msp, offset_x=x_count, offset_y=y_count)
                first_x = x_count
                x_count += 19800000
                y_count += 1960000
                y_offsetcount += 1
                today_count += 1
                second_x = x_count
                flat_sum += (item.get('水平距離', 0) or 0) + (item.get('水平距離2', 0) or 0)
                haikann_sum += (item.get('配管長', 0) or 0) + (item.get('配管長2', 0) or 0)
                msp.add_line((first_x, y_count - y_count - 45000000), (first_x, y_count - y_count - 61000000))
                msp.add_line((first_x + 10000000, y_count - y_count - 45000000), (first_x + 10000000, y_count - y_count - 61000000))
                msp.add_line((second_x, y_count - y_count - 45000000), (second_x, y_count - y_count - 61000000))
                msp.add_line((first_x, y_count - y_count - 61000000), (second_x, y_count - y_count - 61000000))
                cn.flat_line1(msp, item, offset_x=x_count-14200000, offset_y=y_count - y_count - 61000000)
                cn.flat_line2(msp, item, offset_x=x_count-4700000, offset_y=y_count - y_count - 61000000)
                if item['水平距離'] == item['配管長']:
                    pass
                else:
                    cn.haikann_line1(msp, item, offset_x=x_count-19500000, offset_y=y_count - y_count - 54500000)                   
                if item['水平距離2'] == item['配管長2']:
                    pass
                else:
                    cn.haikann_line2(msp, item, offset_x=x_count-12100000, offset_y=y_count - y_count - 54500000)
        elif item['方向'] == 'AB':
            if today_count == 0:
                cn.pe_vb_katauke_bend_11_up_st_ab(msp, offset_x=x_count, offset_y=y_count)
                x_count += 17840000
                y_count += 1570000
                today_count += 1
                today_xbase += 17840000
            else:
                cn.choose_zairyou_no(msp, item, offset_x=x_count + 10000000, offset_y=y_count + 15000000)
                cn.name_sokutenn(msp, item, route_name,offset_x=x_count - 90000000, offset_y=y_count - y_count + 13500000 * y_offsetcount)
                cn.pe_vb_katauke_bend_11_up_st_ab(msp, offset_x=x_count, offset_y=y_count)
                first_x = x_count
                x_count += 17840000
                y_count += 1570000
                y_offsetcount += 1
                today_count += 1
                second_x = x_count
                flat_sum += (item.get('水平距離', 0) or 0) + (item.get('水平距離2', 0) or 0)
                haikann_sum += (item.get('配管長', 0) or 0) + (item.get('配管長2', 0) or 0)
                msp.add_line((first_x, y_count - y_count - 45000000), (first_x, y_count - y_count - 61000000))
                msp.add_line((first_x + 10000000, y_count - y_count - 45000000), (first_x + 10000000, y_count - y_count - 61000000))
                msp.add_line((second_x, y_count - y_count - 45000000), (second_x, y_count - y_count - 61000000))
                msp.add_line((first_x, y_count - y_count - 61000000), (second_x, y_count - y_count - 61000000))
                cn.flat_line1(msp, item, offset_x=x_count-12400000, offset_y=y_count - y_count - 61000000)
                cn.flat_line2(msp, item, offset_x=x_count-3700000, offset_y=y_count - y_count - 61000000)
                if item['水平距離'] == item['配管長']:
                    pass
                else:
                    cn.haikann_line1(msp, item, offset_x=x_count-19500000, offset_y=y_count - y_count - 54500000)                   
                if item['水平距離2'] == item['配管長2']:
                    pass
                else:
                    cn.haikann_line2(msp, item, offset_x=x_count-11100000, offset_y=y_count - y_count - 54500000)
        elif item['方向'] == 'BA':
            if today_count == 0:
                cn.pe_vb_katauke_bend_11_up_st_ba(msp, offset_x=x_count, offset_y=y_count)
                x_count += 17800000
                y_count += 1960000
                today_count += 1
                today_xbase += 17800000
            else:
                cn.choose_zairyou_no(msp, item, offset_x=x_count + 8000000, offset_y=y_count + 15000000)
                cn.name_sokutenn(msp, item, route_name,offset_x=x_count - 90000000, offset_y=y_count - y_count + 13500000 * y_offsetcount)
                cn.pe_vb_katauke_bend_11_up_st_ba(msp, offset_x=x_count, offset_y=y_count)
                first_x = x_count
                x_count += 17800000
                y_count += 1960000
                y_offsetcount += 1
                today_count += 1
                second_x = x_count
                flat_sum += (item.get('水平距離', 0) or 0) + (item.get('水平距離2', 0) or 0)
                haikann_sum += (item.get('配管長', 0) or 0) + (item.get('配管長2', 0) or 0)
                msp.add_line((first_x, y_count - y_count - 45000000), (first_x, y_count - y_count - 61000000))
                msp.add_line((first_x + 8000000, y_count - y_count - 45000000), (first_x + 8000000, y_count - y_count - 61000000))
                msp.add_line((second_x, y_count - y_count - 45000000), (second_x, y_count - y_count - 61000000))
                msp.add_line((first_x, y_count - y_count - 61000000), (second_x, y_count - y_count - 61000000))
                cn.flat_line1(msp, item, offset_x=x_count-12400000, offset_y=y_count - y_count - 61000000)
                cn.flat_line2(msp, item, offset_x=x_count-3700000, offset_y=y_count - y_count - 61000000)
                if item['水平距離'] == item['配管長']:
                    pass
                else:
                    cn.haikann_line1(msp, item, offset_x=x_count-19500000, offset_y=y_count - y_count - 54500000)                   
                if item['水平距離2'] == item['配管長2']:
                    pass
                else:
                    cn.haikann_line2(msp, item, offset_x=x_count-11100000, offset_y=y_count - y_count - 54500000)

    elif (
        'EF' in name
        and 'ベンド' in name
        and '90°' in name
        and '（V）' in name
        and down_90
    ):
        if item['方向'] == 'AA':
            if today_count == 0:
                cn.pe_vb_ryouuke_bend_90_down_ed_aa(msp, offset_x=x_count, offset_y=y_count)
                x_count += 10000000
                y_count += -10000000
                today_count += 1
                today_xbase += 10000000
            else:
                cn.choose_zairyou_no(msp, item, offset_x=x_count + 10000000, offset_y=y_count)
                cn.name_sokutenn(msp, item, route_name,offset_x=x_count - 90000000, offset_y=y_count + 30000000 + 13500000 * y_offsetcount)
                cn.pe_vb_ryouuke_bend_90_down_ed_aa(msp, offset_x=x_count, offset_y=y_count)
                first_x = x_count
                x_count += 10000000
                y_count += -10000000
                y_offsetcount += 1
                today_count += 1
                second_x = x_count
                flat_sum += (item.get('水平距離', 0) or 0) + (item.get('水平距離2', 0) or 0)
                haikann_sum += (item.get('配管長', 0) or 0) + (item.get('配管長2', 0) or 0)
                msp.add_line((first_x, y_count - y_count - 45000000), (first_x, y_count - y_count - 61000000))
                msp.add_line((second_x, y_count - y_count - 45000000), (second_x, y_count - y_count - 61000000))
                msp.add_line((first_x, y_count - y_count - 61000000), (second_x, y_count - y_count - 61000000))
                cn.flat_line2(msp, item, offset_x=x_count - 5000000, offset_y=y_count - y_count - 61000000)
        elif item['方向'] == 'AB':
            if today_count == 0:
                cn.pe_vb_katauke_bend_90_down_ed_ab(msp, offset_x=x_count, offset_y=y_count)
                x_count += 8000000
                y_count += -10000000
                today_count += 1
                today_xbase += 8000000
            else:
                cn.choose_zairyou_no(msp, item, offset_x=x_count + 11000000, offset_y=y_count)
                cn.name_sokutenn(msp, item, route_name,offset_x=x_count - 90000000, offset_y=y_count + 30000000 + 13500000 * y_offsetcount)
                cn.pe_vb_katauke_bend_90_down_ed_ab(msp, offset_x=x_count, offset_y=y_count)
                first_x = x_count
                x_count += 8000000
                y_count += -10000000
                y_offsetcount += 1
                today_count += 1
                second_x = x_count
                flat_sum += (item.get('水平距離', 0) or 0) + (item.get('水平距離2', 0) or 0)
                haikann_sum += (item.get('配管長', 0) or 0) + (item.get('配管長2', 0) or 0)
                msp.add_line((first_x, y_count - y_count - 45000000), (first_x, y_count - y_count - 61000000))
                msp.add_line((second_x, y_count - y_count - 45000000), (second_x, y_count - y_count - 61000000))
                msp.add_line((first_x, y_count - y_count - 61000000), (second_x, y_count - y_count - 61000000))
                cn.flat_line2(msp, item, offset_x=x_count - 4500000, offset_y=y_count - y_count - 61000000)  
        elif item['方向'] == 'BA':
            if today_count == 0:
                cn.pe_vb_katauke_bend_90_down_ed_ba(msp, offset_x=x_count, offset_y=y_count)
                x_count += 10000000
                y_count += -8000000
                today_count += 1
                today_xbase += 10000000
            else:
                cn.choose_zairyou_no(msp, item, offset_x=x_count + 14000000, offset_y=y_count)
                cn.name_sokutenn(msp, item, route_name,offset_x=x_count - 90000000, offset_y=y_count + 30000000 + 13500000 * y_offsetcount)
                cn.pe_vb_katauke_bend_90_down_ed_ba(msp, offset_x=x_count, offset_y=y_count)
                first_x = x_count
                x_count += 10000000
                y_count += -8000000
                y_offsetcount += 1
                today_count += 1
                second_x = x_count
                flat_sum += (item.get('水平距離', 0) or 0) + (item.get('水平距離2', 0) or 0)
                haikann_sum += (item.get('配管長', 0) or 0) + (item.get('配管長2', 0) or 0)
                msp.add_line((first_x, y_count - y_count - 45000000), (first_x, y_count - y_count - 61000000))
                msp.add_line((second_x, y_count - y_count - 45000000), (second_x, y_count - y_count - 61000000))
                msp.add_line((first_x, y_count - y_count - 61000000), (second_x, y_count - y_count - 61000000))
                cn.flat_line2(msp, item, offset_x=x_count - 5000000, offset_y=y_count - y_count - 61000000)                              
    elif (
        'EF' in name
        and 'ベンド' in name
        and '90°' in name
        and '（V）' in name
        and up_90
    ):
        if item['方向'] == 'AA':
            if today_count == 0:
                cn.pe_vb_ryouuke_bend_90_up_ed_aa(msp, offset_x=x_count, offset_y=y_count)
                x_count += 10000000
                y_count += 10000000
                today_count += 1
                today_xbase += 10000000
            else:
                cn.choose_zairyou_no(msp, item, offset_x=x_count -5000000, offset_y=y_count + 15000000)
                cn.name_sokutenn(msp, item, route_name,offset_x=x_count - 90000000, offset_y=y_count + 13500000 * y_offsetcount)
                cn.pe_vb_ryouuke_bend_90_up_ed_aa(msp, offset_x=x_count, offset_y=y_count)
                first_x = x_count
                x_count += 10000000
                y_count += 10000000
                y_offsetcount += 1
                today_count += 1
                second_x = x_count
                flat_sum += (item.get('水平距離', 0) or 0) + (item.get('水平距離2', 0) or 0)
                haikann_sum += (item.get('配管長', 0) or 0) + (item.get('配管長2', 0) or 0)
                msp.add_line((first_x, y_count - y_count - 45000000), (first_x, y_count - y_count - 61000000))
                msp.add_line((second_x, y_count - y_count - 45000000), (second_x, y_count - y_count - 61000000))
                msp.add_line((first_x, y_count - y_count - 61000000), (second_x, y_count - y_count - 61000000))
                cn.flat_line2(msp, item, offset_x=x_count-5000000, offset_y=y_count - y_count - 61000000)
        elif item['方向'] == 'AB':
            if today_count == 0:
                cn.pe_vb_katauke_bend_90_up_ed_ab(msp, offset_x=x_count, offset_y=y_count)
                x_count += 8000000
                y_count += 10000000
                today_count += 1
                today_xbase += 8000000
            else:
                cn.choose_zairyou_no(msp, item, offset_x=x_count -5000000, offset_y=y_count + 15000000)
                cn.name_sokutenn(msp, item, route_name,offset_x=x_count - 90000000, offset_y=y_count + 13500000 * y_offsetcount)
                cn.pe_vb_katauke_bend_90_up_ed_ab(msp, offset_x=x_count, offset_y=y_count)
                first_x = x_count
                x_count += 8000000
                y_count += 10000000
                y_offsetcount += 1
                today_count += 1
                second_x = x_count
                flat_sum += (item.get('水平距離', 0) or 0) + (item.get('水平距離2', 0) or 0)
                haikann_sum += (item.get('配管長', 0) or 0) + (item.get('配管長2', 0) or 0)
                msp.add_line((first_x, y_count - y_count - 45000000), (first_x, y_count - y_count - 61000000))
                msp.add_line((second_x, y_count - y_count - 45000000), (second_x, y_count - y_count - 61000000))
                msp.add_line((first_x, y_count - y_count - 61000000), (second_x, y_count - y_count - 61000000))
                cn.flat_line2(msp, item, offset_x=x_count - 5000000, offset_y=y_count - y_count - 61000000)
        elif item['方向'] == 'BA':
            if today_count == 0:
                cn.pe_vb_katauke_bend_90_up_ed_ba(msp, offset_x=x_count, offset_y=y_count)
                x_count += 10000000
                y_count += 8000000
                today_count += 1
                today_xbase += 10000000
            else:
                cn.choose_zairyou_no(msp, item, offset_x=x_count -5000000, offset_y=y_count + 15000000)
                cn.name_sokutenn(msp, item, route_name,offset_x=x_count - 90000000, offset_y=y_count + 13500000 * y_offsetcount)
                cn.pe_vb_katauke_bend_90_up_ed_ba(msp, offset_x=x_count, offset_y=y_count)
                first_x = x_count
                x_count += 10000000
                y_count += 8000000
                y_offsetcount += 1
                today_count += 1
                second_x = x_count
                flat_sum += (item.get('水平距離', 0) or 0) + (item.get('水平距離2', 0) or 0)
                haikann_sum += (item.get('配管長', 0) or 0) + (item.get('配管長2', 0) or 0)
                msp.add_line((first_x, y_count - y_count - 45000000), (first_x, y_count - y_count - 61000000))
                msp.add_line((second_x, y_count - y_count - 45000000), (second_x, y_count - y_count - 61000000))
                msp.add_line((first_x, y_count - y_count - 61000000), (second_x, y_count - y_count - 61000000))
                cn.flat_line2(msp, item, offset_x=x_count - 5000000, offset_y=y_count - y_count - 61000000)                
    elif (
        'EF' in name
        and 'ベンド' in name
        and '45°' in name
        and '（V）' in name
        and down_45
    ):
        if item['方向'] == 'AA':
            if today_count == 0:
                cn.pe_vb_ryouuke_bend_45_down_ed_aa(msp, offset_x=x_count, offset_y=y_count)
                x_count += 17000000
                y_count += -7000000
                today_count += 1
                today_xbase += 17000000
            else:
                cn.choose_zairyou_no(msp, item, offset_x=x_count + 10000000, offset_y=y_count + 3000000)
                cn.name_sokutenn(msp, item, route_name, offset_x=x_count - 90000000, offset_y=y_count - y_count + 13500000 * y_offsetcount)
                cn.pe_vb_ryouuke_bend_45_down_ed_aa(msp, offset_x=x_count, offset_y=y_count)
                first_x = x_count
                x_count += 17000000
                y_count += -7000000
                y_offsetcount += 1
                today_count += 1
                second_x = x_count
                flat_sum += (item.get('水平距離', 0) or 0) + (item.get('水平距離2', 0) or 0)
                haikann_sum += (item.get('配管長', 0) or 0) + (item.get('配管長2', 0) or 0)
                msp.add_line((first_x, y_count - y_count - 45000000), (first_x, y_count - y_count - 61000000))
                msp.add_line((second_x, y_count - y_count - 45000000), (second_x, y_count - y_count - 61000000))
                msp.add_line((second_x - 10000000, y_count - y_count - 45000000), (second_x - 10000000, y_count - y_count - 61000000))
                msp.add_line((first_x, y_count - y_count - 61000000), (second_x, y_count - y_count - 61000000))
                cn.flat_line1(msp, item, offset_x=x_count-13500000, offset_y=y_count - y_count - 61000000)
                cn.flat_line2(msp, item, offset_x=x_count-5000000, offset_y=y_count - y_count - 61000000)
                if item['水平距離'] == item['配管長']:
                    pass
                else:
                    cn.haikann_line1(msp, item, offset_x=x_count-21000000, offset_y=y_count - y_count - 54500000)                
                if item['水平距離2'] == item['配管長2']:
                    pass
                else:
                    cn.haikann_line2(msp, item, offset_x=x_count-12600000, offset_y=y_count - y_count - 54500000)
        elif item['方向'] == 'AB':
            if today_count == 0:
                cn.pe_vb_katauke_bend_45_down_ed_ab(msp, offset_x=x_count, offset_y=y_count)
                x_count += 15080000
                y_count += -7080000
                today_count += 1
                today_xbase += 15080000
            else:
                cn.choose_zairyou_no(msp, item, offset_x=x_count + 10000000, offset_y=y_count + 3000000)
                cn.name_sokutenn(msp, item, route_name, offset_x=x_count - 90000000, offset_y=y_count - y_count + 13500000 * y_offsetcount)
                cn.pe_vb_katauke_bend_45_down_ed_ab(msp, offset_x=x_count, offset_y=y_count)
                first_x = x_count
                x_count += 15080000
                y_count += -7080000
                y_offsetcount += 1
                today_count += 1
                second_x = x_count
                flat_sum += (item.get('水平距離', 0) or 0) + (item.get('水平距離2', 0) or 0)
                haikann_sum += (item.get('配管長', 0) or 0) + (item.get('配管長2', 0) or 0)
                msp.add_line((first_x, y_count - y_count - 45000000), (first_x, y_count - y_count - 61000000))
                msp.add_line((second_x, y_count - y_count - 45000000), (second_x, y_count - y_count - 61000000))
                msp.add_line((second_x - 10000000, y_count - y_count - 45000000), (second_x - 10000000, y_count - y_count - 61000000))
                msp.add_line((first_x, y_count - y_count - 61000000), (second_x, y_count - y_count - 61000000))
                cn.flat_line1(msp, item, offset_x=x_count-13500000, offset_y=y_count - y_count - 61000000)
                cn.flat_line2(msp, item, offset_x=x_count-5000000, offset_y=y_count - y_count - 61000000)
                if item['水平距離'] == item['配管長']:
                    pass
                else:
                    cn.haikann_line1(msp, item, offset_x=x_count-21000000, offset_y=y_count - y_count - 54500000)                
                if item['水平距離2'] == item['配管長2']:
                    pass
                else:
                    cn.haikann_line2(msp, item, offset_x=x_count-12600000, offset_y=y_count - y_count - 54500000)
        elif item['方向'] == 'BA':
            if today_count == 0:
                cn.pe_vb_katauke_bend_45_down_ed_ba(msp, offset_x=x_count, offset_y=y_count)
                x_count += 15660000
                y_count += -5660000
                today_count += 1
                today_xbase += 15660000
            else:
                cn.choose_zairyou_no(msp, item, offset_x=x_count + 10000000, offset_y=y_count + 3000000)
                cn.name_sokutenn(msp, item, route_name, offset_x=x_count - 90000000, offset_y=y_count - y_count + 13500000 * y_offsetcount)
                cn.pe_vb_katauke_bend_45_down_ed_ba(msp, offset_x=x_count, offset_y=y_count)
                first_x = x_count
                x_count += 15660000
                y_count += -5660000
                y_offsetcount += 1
                today_count += 1
                second_x = x_count
                flat_sum += (item.get('水平距離', 0) or 0) + (item.get('水平距離2', 0) or 0)
                haikann_sum += (item.get('配管長', 0) or 0) + (item.get('配管長2', 0) or 0)
                msp.add_line((first_x, y_count - y_count - 45000000), (first_x, y_count - y_count - 61000000))
                msp.add_line((second_x, y_count - y_count - 45000000), (second_x, y_count - y_count - 61000000))
                msp.add_line((second_x - 10000000, y_count - y_count - 45000000), (second_x - 10000000, y_count - y_count - 61000000))
                msp.add_line((first_x, y_count - y_count - 61000000), (second_x, y_count - y_count - 61000000))
                cn.flat_line1(msp, item, offset_x=x_count-13500000, offset_y=y_count - y_count - 61000000)
                cn.flat_line2(msp, item, offset_x=x_count-5000000, offset_y=y_count - y_count - 61000000)
                if item['水平距離'] == item['配管長']:
                    pass
                else:
                    cn.haikann_line1(msp, item, offset_x=x_count-21000000, offset_y=y_count - y_count - 54500000)                
                if item['水平距離2'] == item['配管長2']:
                    pass
                else:
                    cn.haikann_line2(msp, item, offset_x=x_count-12600000, offset_y=y_count - y_count - 54500000)
    elif (
        'EF' in name
        and 'ベンド' in name
        and '45°' in name
        and '（V）' in name
        and up_45
    ):
        if item['方向'] == 'AA':
            if today_count == 0:
                cn.pe_vb_ryouuke_bend_45_up_ed_aa(msp, offset_x=x_count, offset_y=y_count)
                x_count += 17000000
                y_count += 7000000
                today_count += 1
                today_xbase += 17000000
            else:
                cn.choose_zairyou_no(msp, item, offset_x=x_count + 8500000, offset_y=y_count + 15000000)
                cn.name_sokutenn(msp, item, route_name, offset_x=x_count - 90000000, offset_y=y_count - y_count + 13500000 * y_offsetcount)
                cn.pe_vb_ryouuke_bend_45_up_ed_aa(msp, offset_x=x_count, offset_y=y_count)
                first_x = x_count
                x_count += 17000000
                y_count += 7000000
                y_offsetcount += 1
                today_count += 1
                second_x = x_count
                flat_sum += (item.get('水平距離', 0) or 0) + (item.get('水平距離2', 0) or 0)
                haikann_sum += (item.get('配管長', 0) or 0) + (item.get('配管長2', 0) or 0)
                msp.add_line((first_x, y_count - y_count - 45000000), (first_x, y_count - y_count - 61000000))
                msp.add_line((second_x, y_count - y_count - 45000000), (second_x, y_count - y_count - 61000000))
                msp.add_line((second_x - 10000000, y_count - y_count - 45000000), (second_x - 10000000, y_count - y_count - 61000000))
                msp.add_line((first_x, y_count - y_count - 61000000), (second_x, y_count - y_count - 61000000))
                cn.flat_line1(msp, item, offset_x=x_count-13500000, offset_y=y_count - y_count - 61000000)
                cn.flat_line2(msp, item, offset_x=x_count-5000000, offset_y=y_count - y_count - 61000000)
                if item['水平距離'] == item['配管長']:
                    pass
                else:
                    cn.haikann_line1(msp, item, offset_x=x_count-21000000, offset_y=y_count - y_count - 54500000)                
                if item['水平距離2'] == item['配管長2']:
                    pass
                else:
                    cn.haikann_line2(msp, item, offset_x=x_count-12600000, offset_y=y_count - y_count - 54500000)
        elif item['方向'] == 'AB':
            if today_count == 0:
                cn.pe_vb_katauke_bend_45_up_ed_ab(msp, offset_x=x_count, offset_y=y_count)
                x_count += 15080000
                y_count += 7080000
                today_count += 1
                today_xbase += 15080000
            else:
                cn.choose_zairyou_no(msp, item, offset_x=x_count + 8500000, offset_y=y_count + 15000000)
                cn.name_sokutenn(msp, item, route_name, offset_x=x_count - 90000000, offset_y=y_count - y_count + 13500000 * y_offsetcount)
                cn.pe_vb_katauke_bend_45_up_ed_ab(msp, offset_x=x_count, offset_y=y_count)
                first_x = x_count
                x_count += 15080000
                y_count += 7080000
                y_offsetcount += 1
                today_count += 1
                second_x = x_count
                flat_sum += (item.get('水平距離', 0) or 0) + (item.get('水平距離2', 0) or 0)
                haikann_sum += (item.get('配管長', 0) or 0) + (item.get('配管長2', 0) or 0)
                msp.add_line((first_x, y_count - y_count - 45000000), (first_x, y_count - y_count - 61000000))
                msp.add_line((second_x, y_count - y_count - 45000000), (second_x, y_count - y_count - 61000000))
                msp.add_line((second_x - 10000000, y_count - y_count - 45000000), (second_x - 10000000, y_count - y_count - 61000000))
                msp.add_line((first_x, y_count - y_count - 61000000), (second_x, y_count - y_count - 61000000))
                cn.flat_line1(msp, item, offset_x=x_count-13500000, offset_y=y_count - y_count - 61000000)
                cn.flat_line2(msp, item, offset_x=x_count-5000000, offset_y=y_count - y_count - 61000000)
                if item['水平距離'] == item['配管長']:
                    pass
                else:
                    cn.haikann_line1(msp, item, offset_x=x_count-21000000, offset_y=y_count - y_count - 54500000)                
                if item['水平距離2'] == item['配管長2']:
                    pass
                else:
                    cn.haikann_line2(msp, item, offset_x=x_count-12600000, offset_y=y_count - y_count - 54500000)
        elif item['方向'] == 'BA':
            if today_count == 0:
                cn.pe_vb_katauke_bend_45_up_ed_ba(msp, offset_x=x_count, offset_y=y_count)
                x_count += 15660000
                y_count += 5660000
                today_count += 1
                today_xbase += 15660000
            else:
                cn.choose_zairyou_no(msp, item, offset_x=x_count + 8500000, offset_y=y_count + 15000000)
                cn.name_sokutenn(msp, item, route_name, offset_x=x_count - 90000000, offset_y=y_count - y_count + 13500000 * y_offsetcount)
                cn.pe_vb_katauke_bend_45_up_ed_ba(msp, offset_x=x_count, offset_y=y_count)
                first_x = x_count
                x_count += 15660000
                y_count += 5660000
                y_offsetcount += 1
                today_count += 1
                second_x = x_count
                flat_sum += (item.get('水平距離', 0) or 0) + (item.get('水平距離2', 0) or 0)
                haikann_sum += (item.get('配管長', 0) or 0) + (item.get('配管長2', 0) or 0)
                msp.add_line((first_x, y_count - y_count - 45000000), (first_x, y_count - y_count - 61000000))
                msp.add_line((second_x, y_count - y_count - 45000000), (second_x, y_count - y_count - 61000000))
                msp.add_line((second_x - 10000000, y_count - y_count - 45000000), (second_x - 10000000, y_count - y_count - 61000000))
                msp.add_line((first_x, y_count - y_count - 61000000), (second_x, y_count - y_count - 61000000))
                cn.flat_line1(msp, item, offset_x=x_count-13500000, offset_y=y_count - y_count - 61000000)
                cn.flat_line2(msp, item, offset_x=x_count-5000000, offset_y=y_count - y_count - 61000000)
                if item['水平距離'] == item['配管長']:
                    pass
                else:
                    cn.haikann_line1(msp, item, offset_x=x_count-21000000, offset_y=y_count - y_count - 54500000)                
                if item['水平距離2'] == item['配管長2']:
                    pass
                else:
                    cn.haikann_line2(msp, item, offset_x=x_count-12600000, offset_y=y_count - y_count - 54500000)
    elif (
        'EF' in name
        and 'ベンド' in name
        and '22' in name
        and '（V）' in name
        and down_22
    ):
        if item['方向'] == 'AA':
            if today_count == 0:
                cn.pe_vb_ryouuke_bend_22_down_ed_aa(msp, offset_x=x_count, offset_y=y_count)
                x_count += 19240000
                y_count += -3820000
                today_count += 1
                today_xbase += 19240000
            else:
                cn.choose_zairyou_no(msp, item, offset_x=x_count + 9240000, offset_y=y_count + 11180000)
                cn.name_sokutenn(msp, item, route_name, offset_x=x_count - 90000000, offset_y=y_count - y_count + 13500000 * y_offsetcount)
                cn.pe_vb_ryouuke_bend_22_down_ed_aa(msp, offset_x=x_count, offset_y=y_count)
                first_x = x_count
                x_count += 19240000
                y_count += -3820000
                y_offsetcount += 1
                today_count += 1
                second_x = x_count
                flat_sum += (item.get('水平距離', 0) or 0) + (item.get('水平距離2', 0) or 0)
                haikann_sum += (item.get('配管長', 0) or 0) + (item.get('配管長2', 0) or 0)
                msp.add_line((first_x, y_count - y_count - 45000000), (first_x, y_count - y_count - 61000000))
                msp.add_line((second_x, y_count - y_count - 45000000), (second_x, y_count - y_count - 61000000))
                msp.add_line((second_x - 10000000, y_count - y_count - 45000000), (second_x - 10000000, y_count - y_count - 61000000))
                msp.add_line((first_x, y_count - y_count - 61000000), (second_x, y_count - y_count - 61000000))
                cn.flat_line1(msp, item, offset_x=x_count-14650000, offset_y=y_count - y_count - 61000000)
                cn.flat_line2(msp, item, offset_x=x_count-5000000, offset_y=y_count - y_count - 61000000)
                if item['水平距離'] == item['配管長']:
                    pass
                else:
                    cn.haikann_line1(msp, item, offset_x=x_count-22150000, offset_y=y_count - y_count - 54500000)                
                if item['水平距離2'] == item['配管長2']:
                    pass
                else:
                    cn.haikann_line2(msp, item, offset_x=x_count-12600000, offset_y=y_count - y_count - 54500000)
        elif item['方向'] == 'AB':
            if today_count == 0:
                cn.pe_vb_katauke_bend_22_down_ed_ab(msp, offset_x=x_count, offset_y=y_count)
                x_count += 17240000
                y_count += -3820000
                today_count += 1
                today_xbase += 17240000
            else:
                cn.choose_zairyou_no(msp, item, offset_x=x_count + 8620000, offset_y=y_count + 11180000)
                cn.name_sokutenn(msp, item, route_name, offset_x=x_count - 90000000, offset_y=y_count - y_count + 13500000 * y_offsetcount)
                cn.pe_vb_katauke_bend_22_down_ed_ab(msp, offset_x=x_count, offset_y=y_count)
                first_x = x_count
                x_count += 17240000
                y_count += -3820000
                y_offsetcount += 1
                today_count += 1
                second_x = x_count
                flat_sum += (item.get('水平距離', 0) or 0) + (item.get('水平距離2', 0) or 0)
                haikann_sum += (item.get('配管長', 0) or 0) + (item.get('配管長2', 0) or 0)
                msp.add_line((first_x, y_count - y_count - 45000000), (first_x, y_count - y_count - 61000000))
                msp.add_line((second_x, y_count - y_count - 45000000), (second_x, y_count - y_count - 61000000))
                msp.add_line((second_x - 10000000, y_count - y_count - 45000000), (second_x - 10000000, y_count - y_count - 61000000))
                msp.add_line((first_x, y_count - y_count - 61000000), (second_x, y_count - y_count - 61000000))
                cn.flat_line1(msp, item, offset_x=x_count-13600000, offset_y=y_count - y_count - 61000000)
                cn.flat_line2(msp, item, offset_x=x_count-5000000, offset_y=y_count - y_count - 61000000)
                if item['水平距離'] == item['配管長']:
                    pass
                else:
                    cn.haikann_line1(msp, item, offset_x=x_count-21100000, offset_y=y_count - y_count - 54500000)                
                if item['水平距離2'] == item['配管長2']:
                    pass
                else:
                    cn.haikann_line2(msp, item, offset_x=x_count-12600000, offset_y=y_count - y_count - 54500000)
        elif item['方向'] == 'BA':
            if today_count == 0:
                cn.pe_vb_katauke_bend_22_down_ed_ba(msp, offset_x=x_count, offset_y=y_count)
                x_count += 17240000
                y_count += -3820000
                today_count += 1
                today_xbase += 17240000
            else:
                cn.choose_zairyou_no(msp, item, offset_x=x_count + 8620000, offset_y=y_count + 11180000)
                cn.name_sokutenn(msp, item, route_name, offset_x=x_count - 90000000, offset_y=y_count - y_count + 13500000 * y_offsetcount)
                cn.pe_vb_katauke_bend_22_down_ed_ba(msp, offset_x=x_count, offset_y=y_count)
                first_x = x_count
                x_count += 17240000
                y_count += -3820000
                y_offsetcount += 1
                today_count += 1
                second_x = x_count
                flat_sum += (item.get('水平距離', 0) or 0) + (item.get('水平距離2', 0) or 0)
                haikann_sum += (item.get('配管長', 0) or 0) + (item.get('配管長2', 0) or 0)
                msp.add_line((first_x, y_count - y_count - 45000000), (first_x, y_count - y_count - 61000000))
                msp.add_line((second_x, y_count - y_count - 45000000), (second_x, y_count - y_count - 61000000))
                msp.add_line((second_x - 10000000, y_count - y_count - 45000000), (second_x - 10000000, y_count - y_count - 61000000))
                msp.add_line((first_x, y_count - y_count - 61000000), (second_x, y_count - y_count - 61000000))
                cn.flat_line1(msp, item, offset_x=x_count-13600000, offset_y=y_count - y_count - 61000000)
                cn.flat_line2(msp, item, offset_x=x_count-5000000, offset_y=y_count - y_count - 61000000)
                if item['水平距離'] == item['配管長']:
                    pass
                else:
                    cn.haikann_line1(msp, item, offset_x=x_count-21100000, offset_y=y_count - y_count - 54500000)                
                if item['水平距離2'] == item['配管長2']:
                    pass
                else:
                    cn.haikann_line2(msp, item, offset_x=x_count-12600000, offset_y=y_count - y_count - 54500000)

    elif (
        'EF' in name
        and 'ベンド' in name
        and '22' in name
        and '（V）' in name
        and up_22
    ):
        if item['方向'] == 'AA':
            if today_count == 0:
                cn.pe_vb_ryouuke_bend_22_up_ed_aa(msp, offset_x=x_count, offset_y=y_count)
                x_count += 19240000
                y_count += 3820000
                today_count += 1
                today_xbase += 19240000
            else:
                cn.choose_zairyou_no(msp, item, offset_x=x_count + 9240000, offset_y=y_count + 18820000)
                cn.name_sokutenn(msp, item, route_name, offset_x=x_count - 90000000, offset_y=y_count - y_count + 13500000 * y_offsetcount)
                cn.pe_vb_ryouuke_bend_22_up_ed_aa(msp, offset_x=x_count, offset_y=y_count)
                first_x = x_count
                x_count += 19240000
                y_count += 3820000
                y_offsetcount += 1
                today_count += 1
                second_x = x_count
                flat_sum += (item.get('水平距離', 0) or 0) + (item.get('水平距離2', 0) or 0)
                haikann_sum += (item.get('配管長', 0) or 0) + (item.get('配管長2', 0) or 0)
                msp.add_line((first_x, y_count - y_count - 45000000), (first_x, y_count - y_count - 61000000))
                msp.add_line((second_x, y_count - y_count - 45000000), (second_x, y_count - y_count - 61000000))
                msp.add_line((second_x - 10000000, y_count - y_count - 45000000), (second_x - 10000000, y_count - y_count - 61000000))
                msp.add_line((first_x, y_count - y_count - 61000000), (second_x, y_count - y_count - 61000000))
                cn.flat_line1(msp, item, offset_x=x_count-14650000, offset_y=y_count - y_count - 61000000)
                cn.flat_line2(msp, item, offset_x=x_count-5000000, offset_y=y_count - y_count - 61000000)
                if item['水平距離'] == item['配管長']:
                    pass
                else:
                    cn.haikann_line1(msp, item, offset_x=x_count-22150000, offset_y=y_count - y_count - 54500000)                
                if item['水平距離2'] == item['配管長2']:
                    pass
                else:
                    cn.haikann_line2(msp, item, offset_x=x_count-12600000, offset_y=y_count - y_count - 54500000)
        elif item['方向'] == 'AB':
            if today_count == 0:
                cn.pe_vb_katauke_bend_22_up_ed_ab(msp, offset_x=x_count, offset_y=y_count)
                x_count += 17240000
                y_count += 3820000
                today_count += 1
                today_xbase += 17240000
            else:
                cn.choose_zairyou_no(msp, item, offset_x=x_count + 9240000, offset_y=y_count + 18820000)
                cn.name_sokutenn(msp, item, route_name, offset_x=x_count - 90000000, offset_y=y_count - y_count + 13500000 * y_offsetcount)
                cn.pe_vb_katauke_bend_22_up_ed_ab(msp, offset_x=x_count, offset_y=y_count)
                first_x = x_count
                x_count += 17240000
                y_count += 3820000
                y_offsetcount += 1
                today_count += 1
                second_x = x_count
                flat_sum += (item.get('水平距離', 0) or 0) + (item.get('水平距離2', 0) or 0)
                haikann_sum += (item.get('配管長', 0) or 0) + (item.get('配管長2', 0) or 0)
                msp.add_line((first_x, y_count - y_count - 45000000), (first_x, y_count - y_count - 61000000))
                msp.add_line((second_x, y_count - y_count - 45000000), (second_x, y_count - y_count - 61000000))
                msp.add_line((second_x - 10000000, y_count - y_count - 45000000), (second_x - 10000000, y_count - y_count - 61000000))
                msp.add_line((first_x, y_count - y_count - 61000000), (second_x, y_count - y_count - 61000000))
                cn.flat_line1(msp, item, offset_x=x_count-13500000, offset_y=y_count - y_count - 61000000)
                cn.flat_line2(msp, item, offset_x=x_count-5000000, offset_y=y_count - y_count - 61000000)
                if item['水平距離'] == item['配管長']:
                    pass
                else:
                    cn.haikann_line1(msp, item, offset_x=x_count-21000000, offset_y=y_count - y_count - 54500000)                
                if item['水平距離2'] == item['配管長2']:
                    pass
                else:
                    cn.haikann_line2(msp, item, offset_x=x_count-12600000, offset_y=y_count - y_count - 54500000)
        elif item['方向'] == 'BA':
            if today_count == 0:
                cn.pe_vb_katauke_bend_22_up_ed_ba(msp, offset_x=x_count, offset_y=y_count)
                x_count += 17380000
                y_count += 3060000
                today_count += 1
                today_xbase += 17380000
            else:
                cn.choose_zairyou_no(msp, item, offset_x=x_count + 7380000, offset_y=y_count + 18060000)
                cn.name_sokutenn(msp, item, route_name, offset_x=x_count - 90000000, offset_y=y_count - y_count + 13500000 * y_offsetcount)
                cn.pe_vb_katauke_bend_22_up_ed_ba(msp, offset_x=x_count, offset_y=y_count)
                first_x = x_count
                x_count += 17380000
                y_count += 3060000
                y_offsetcount += 1
                today_count += 1
                second_x = x_count
                flat_sum += (item.get('水平距離', 0) or 0) + (item.get('水平距離2', 0) or 0)
                haikann_sum += (item.get('配管長', 0) or 0) + (item.get('配管長2', 0) or 0)
                msp.add_line((first_x, y_count - y_count - 45000000), (first_x, y_count - y_count - 61000000))
                msp.add_line((second_x, y_count - y_count - 45000000), (second_x, y_count - y_count - 61000000))
                msp.add_line((second_x - 10000000, y_count - y_count - 45000000), (second_x - 10000000, y_count - y_count - 61000000))
                msp.add_line((first_x, y_count - y_count - 61000000), (second_x, y_count - y_count - 61000000))
                cn.flat_line1(msp, item, offset_x=x_count-13500000, offset_y=y_count - y_count - 61000000)
                cn.flat_line2(msp, item, offset_x=x_count-5000000, offset_y=y_count - y_count - 61000000)
                if item['水平距離'] == item['配管長']:
                    pass
                else:
                    cn.haikann_line1(msp, item, offset_x=x_count-21000000, offset_y=y_count - y_count - 54500000)                
                if item['水平距離2'] == item['配管長2']:
                    pass
                else:
                    cn.haikann_line2(msp, item, offset_x=x_count-12600000, offset_y=y_count - y_count - 54500000)

    elif (
        'EF' in name
        and 'ベンド' in name
        and '11' in name
        and '（V）' in name
        and down_11
    ):
        if item['方向'] == 'AA':
            if today_count == 0:
                cn.pe_vb_ryouuke_bend_11_down_ed_aa(msp, offset_x=x_count, offset_y=y_count)
                x_count += 19800000
                y_count += -1960000
                today_count += 1
                today_xbase += 19800000
            else:
                cn.choose_zairyou_no(msp, item, offset_x=x_count + 9800000, offset_y=y_count + 11180000)
                cn.name_sokutenn(msp, item, route_name, offset_x=x_count - 90000000, offset_y=y_count - y_count + 13500000 * y_offsetcount)
                cn.pe_vb_ryouuke_bend_11_down_ed_aa(msp, offset_x=x_count, offset_y=y_count)
                first_x = x_count
                x_count += 19800000
                y_count += -1960000
                y_offsetcount += 1
                today_count += 1
                second_x = x_count
                flat_sum += (item.get('水平距離', 0) or 0) + (item.get('水平距離2', 0) or 0)
                haikann_sum += (item.get('配管長', 0) or 0) + (item.get('配管長2', 0) or 0)
                msp.add_line((first_x, y_count - y_count - 45000000), (first_x, y_count - y_count - 61000000))
                msp.add_line((second_x, y_count - y_count - 45000000), (second_x, y_count - y_count - 61000000))
                msp.add_line((second_x - 9800000, y_count - y_count - 45000000), (second_x - 9800000, y_count - y_count - 61000000))
                msp.add_line((first_x, y_count - y_count - 61000000), (second_x, y_count - y_count - 61000000))
                cn.flat_line1(msp, item, offset_x=x_count-14650000, offset_y=y_count - y_count - 61000000)
                cn.flat_line2(msp, item, offset_x=x_count-5000000, offset_y=y_count - y_count - 61000000)
                if item['水平距離'] == item['配管長']:
                    pass
                else:
                    cn.haikann_line1(msp, item, offset_x=x_count-22150000, offset_y=y_count - y_count - 54500000)                
                if item['水平距離2'] == item['配管長2']:
                    pass
                else:
                    cn.haikann_line2(msp, item, offset_x=x_count-12600000, offset_y=y_count - y_count - 54500000)
        elif item['方向'] == 'AB':
            if today_count == 0:
                cn.pe_vb_katauke_bend_11_down_ed_ab(msp, offset_x=x_count, offset_y=y_count)
                x_count += 17800000
                y_count += -1960000
                today_count += 1
                today_xbase += 17800000
            else:
                cn.choose_zairyou_no(msp, item, offset_x=x_count + 9800000, offset_y=y_count + 11180000)
                cn.name_sokutenn(msp, item, route_name, offset_x=x_count - 90000000, offset_y=y_count - y_count + 13500000 * y_offsetcount)
                cn.pe_vb_katauke_bend_11_down_ed_ab(msp, offset_x=x_count, offset_y=y_count)
                first_x = x_count
                x_count += 17800000
                y_count += -1960000
                y_offsetcount += 1
                today_count += 1
                second_x = x_count
                flat_sum += (item.get('水平距離', 0) or 0) + (item.get('水平距離2', 0) or 0)
                haikann_sum += (item.get('配管長', 0) or 0) + (item.get('配管長2', 0) or 0)
                msp.add_line((first_x, y_count - y_count - 45000000), (first_x, y_count - y_count - 61000000))
                msp.add_line((second_x, y_count - y_count - 45000000), (second_x, y_count - y_count - 61000000))
                msp.add_line((second_x - 9800000, y_count - y_count - 45000000), (second_x - 9800000, y_count - y_count - 61000000))
                msp.add_line((first_x, y_count - y_count - 61000000), (second_x, y_count - y_count - 61000000))
                cn.flat_line1(msp, item, offset_x=x_count-13600000, offset_y=y_count - y_count - 61000000)
                cn.flat_line2(msp, item, offset_x=x_count-5000000, offset_y=y_count - y_count - 61000000)
                if item['水平距離'] == item['配管長']:
                    pass
                else:
                    cn.haikann_line1(msp, item, offset_x=x_count-21100000, offset_y=y_count - y_count - 54500000)                
                if item['水平距離2'] == item['配管長2']:
                    pass
                else:
                    cn.haikann_line2(msp, item, offset_x=x_count-12600000, offset_y=y_count - y_count - 54500000)
        elif item['方向'] == 'BA':
            if today_count == 0:
                cn.pe_vb_katauke_bend_11_down_ed_ba(msp, offset_x=x_count, offset_y=y_count)
                x_count += 17840000
                y_count += -1570000
                today_count += 1
                today_xbase += 17840000
            else:
                cn.choose_zairyou_no(msp, item, offset_x=x_count + 7840000, offset_y=y_count + 11180000)
                cn.name_sokutenn(msp, item, route_name, offset_x=x_count - 90000000, offset_y=y_count - y_count + 13500000 * y_offsetcount)
                cn.pe_vb_katauke_bend_11_down_ed_ba(msp, offset_x=x_count, offset_y=y_count)
                first_x = x_count
                x_count += 17840000
                y_count += -1570000
                y_offsetcount += 1
                today_count += 1
                second_x = x_count
                flat_sum += (item.get('水平距離', 0) or 0) + (item.get('水平距離2', 0) or 0)
                haikann_sum += (item.get('配管長', 0) or 0) + (item.get('配管長2', 0) or 0)
                msp.add_line((first_x, y_count - y_count - 45000000), (first_x, y_count - y_count - 61000000))
                msp.add_line((second_x, y_count - y_count - 45000000), (second_x, y_count - y_count - 61000000))
                msp.add_line((second_x - 7840000, y_count - y_count - 45000000), (second_x - 7840000, y_count - y_count - 61000000))
                msp.add_line((first_x, y_count - y_count - 61000000), (second_x, y_count - y_count - 61000000))
                cn.flat_line1(msp, item, offset_x=x_count-13600000, offset_y=y_count - y_count - 61000000)
                cn.flat_line2(msp, item, offset_x=x_count-5000000, offset_y=y_count - y_count - 61000000)
                if item['水平距離'] == item['配管長']:
                    pass
                else:
                    cn.haikann_line1(msp, item, offset_x=x_count-21100000, offset_y=y_count - y_count - 54500000)                
                if item['水平距離2'] == item['配管長2']:
                    pass
                else:
                    cn.haikann_line2(msp, item, offset_x=x_count-12600000, offset_y=y_count - y_count - 54500000)

    elif (
        'EF' in name
        and 'ベンド' in name
        and '11' in name
        and '（V）' in name
        and up_11
    ):
        if item['方向'] == 'AA':
            if today_count == 0:
                cn.pe_vb_ryouuke_bend_11_up_ed_aa(msp, offset_x=x_count, offset_y=y_count)
                x_count += 19800000
                y_count += 1960000
                today_count += 1
                today_xbase += 19800000
            else:
                cn.choose_zairyou_no(msp, item, offset_x=x_count + 9800000, offset_y=y_count + 18820000)
                cn.name_sokutenn(msp, item, route_name, offset_x=x_count - 90000000, offset_y=y_count - y_count + 13500000 * y_offsetcount)
                cn.pe_vb_ryouuke_bend_11_up_ed_aa(msp, offset_x=x_count, offset_y=y_count)
                first_x = x_count
                x_count += 19800000
                y_count += 1960000
                y_offsetcount += 1
                today_count += 1
                second_x = x_count
                flat_sum += (item.get('水平距離', 0) or 0) + (item.get('水平距離2', 0) or 0)
                haikann_sum += (item.get('配管長', 0) or 0) + (item.get('配管長2', 0) or 0)
                msp.add_line((first_x, y_count - y_count - 45000000), (first_x, y_count - y_count - 61000000))
                msp.add_line((second_x, y_count - y_count - 45000000), (second_x, y_count - y_count - 61000000))
                msp.add_line((second_x - 9800000, y_count - y_count - 45000000), (second_x - 9800000, y_count - y_count - 61000000))
                msp.add_line((first_x, y_count - y_count - 61000000), (second_x, y_count - y_count - 61000000))
                cn.flat_line1(msp, item, offset_x=x_count-14650000, offset_y=y_count - y_count - 61000000)
                cn.flat_line2(msp, item, offset_x=x_count-5000000, offset_y=y_count - y_count - 61000000)
                if item['水平距離'] == item['配管長']:
                    pass
                else:
                    cn.haikann_line1(msp, item, offset_x=x_count-22150000, offset_y=y_count - y_count - 54500000)                
                if item['水平距離2'] == item['配管長2']:
                    pass
                else:
                    cn.haikann_line2(msp, item, offset_x=x_count-12600000, offset_y=y_count - y_count - 54500000)
        elif item['方向'] == 'AB':
            if today_count == 0:
                cn.pe_vb_katauke_bend_11_up_ed_ab(msp, offset_x=x_count, offset_y=y_count)
                x_count += 17800000
                y_count += 1960000
                today_count += 1
                today_xbase += 17800000
            else:
                cn.choose_zairyou_no(msp, item, offset_x=x_count + 9800000, offset_y=y_count + 18820000)
                cn.name_sokutenn(msp, item, route_name, offset_x=x_count - 90000000, offset_y=y_count - y_count + 13500000 * y_offsetcount)
                cn.pe_vb_katauke_bend_11_up_ed_ab(msp, offset_x=x_count, offset_y=y_count)
                first_x = x_count
                x_count += 17800000
                y_count += 1960000
                y_offsetcount += 1
                today_count += 1
                second_x = x_count
                flat_sum += (item.get('水平距離', 0) or 0) + (item.get('水平距離2', 0) or 0)
                haikann_sum += (item.get('配管長', 0) or 0) + (item.get('配管長2', 0) or 0)
                msp.add_line((first_x, y_count - y_count - 45000000), (first_x, y_count - y_count - 61000000))
                msp.add_line((second_x, y_count - y_count - 45000000), (second_x, y_count - y_count - 61000000))
                msp.add_line((second_x - 9800000, y_count - y_count - 45000000), (second_x - 9800000, y_count - y_count - 61000000))
                msp.add_line((first_x, y_count - y_count - 61000000), (second_x, y_count - y_count - 61000000))
                cn.flat_line1(msp, item, offset_x=x_count-13500000, offset_y=y_count - y_count - 61000000)
                cn.flat_line2(msp, item, offset_x=x_count-5000000, offset_y=y_count - y_count - 61000000)
                if item['水平距離'] == item['配管長']:
                    pass
                else:
                    cn.haikann_line1(msp, item, offset_x=x_count-21000000, offset_y=y_count - y_count - 54500000)                
                if item['水平距離2'] == item['配管長2']:
                    pass
                else:
                    cn.haikann_line2(msp, item, offset_x=x_count-12600000, offset_y=y_count - y_count - 54500000)
        elif item['方向'] == 'BA':
            if today_count == 0:
                cn.pe_vb_katauke_bend_11_up_ed_ba(msp, offset_x=x_count, offset_y=y_count)
                x_count += 17840000
                y_count += 1570000
                today_count += 1
                today_xbase += 17840000
            else:
                cn.choose_zairyou_no(msp, item, offset_x=x_count + 7840000, offset_y=y_count + 18060000)
                cn.name_sokutenn(msp, item, route_name, offset_x=x_count - 90000000, offset_y=y_count - y_count + 13500000 * y_offsetcount)
                cn.pe_vb_katauke_bend_11_up_ed_ba(msp, offset_x=x_count, offset_y=y_count)
                first_x = x_count
                x_count += 17840000
                y_count += 1570000
                y_offsetcount += 1
                today_count += 1
                second_x = x_count
                flat_sum += (item.get('水平距離', 0) or 0) + (item.get('水平距離2', 0) or 0)
                haikann_sum += (item.get('配管長', 0) or 0) + (item.get('配管長2', 0) or 0)
                msp.add_line((first_x, y_count - y_count - 45000000), (first_x, y_count - y_count - 61000000))
                msp.add_line((second_x, y_count - y_count - 45000000), (second_x, y_count - y_count - 61000000))
                msp.add_line((second_x - 7840000, y_count - y_count - 45000000), (second_x - 7840000, y_count - y_count - 61000000))
                msp.add_line((first_x, y_count - y_count - 61000000), (second_x, y_count - y_count - 61000000))
                cn.flat_line1(msp, item, offset_x=x_count-13500000, offset_y=y_count - y_count - 61000000)
                cn.flat_line2(msp, item, offset_x=x_count-5000000, offset_y=y_count - y_count - 61000000)
                if item['水平距離'] == item['配管長']:
                    pass
                else:
                    cn.haikann_line1(msp, item, offset_x=x_count-21000000, offset_y=y_count - y_count - 54500000)                
                if item['水平距離2'] == item['配管長2']:
                    pass
                else:
                    cn.haikann_line2(msp, item, offset_x=x_count-12600000, offset_y=y_count - y_count - 54500000)

    elif (
        'EF' in name
        and 'ベンド' in name
    ):
        if item['方向'] == 'AA':
            if today_count == 0:
                cn.pe_hb_ryouuke_bend_aa(msp, offset_x=x_count, offset_y=y_count)
                x_count += 20000000
                y_offsetcount += 1
                today_count += 1
                today_xbase += 20000000
            else:
                cn.choose_zairyou_no(msp, item, offset_x=x_count + 10000000, offset_y=y_count + 15000000)
                cn.name_sokutenn(msp, item, route_name, offset_x=x_count - 90000000, offset_y=y_count - y_count + 13500000 * y_offsetcount)
                cn.pe_hb_ryouuke_bend_aa(msp, offset_x=x_count, offset_y=y_count)
                first_x = x_count
                x_count += 20000000
                y_offsetcount += 1
                today_count += 1
                second_x = x_count
                flat_sum += (item.get('水平距離', 0) or 0) + (item.get('水平距離2', 0) or 0)
                haikann_sum += (item.get('配管長', 0) or 0) + (item.get('配管長2', 0) or 0)
                msp.add_line((first_x, y_count - y_count - 45000000), (first_x, y_count - y_count - 61000000))
                msp.add_line((second_x, y_count - y_count - 45000000), (second_x, y_count - y_count - 61000000))
                msp.add_line((second_x - 10000000, y_count - y_count - 45000000), (second_x - 10000000, y_count - y_count - 61000000))
                msp.add_line((first_x, y_count - y_count - 61000000), (second_x, y_count - y_count - 61000000))
                cn.flat_line1(msp, item, offset_x=x_count-15000000, offset_y=y_count - y_count - 61000000)
                cn.flat_line2(msp, item, offset_x=x_count-5000000, offset_y=y_count - y_count - 61000000)

        elif item['方向'] == 'AB':
            if today_count == 0:
                cn.pe_hb_katauke_bend_ab(msp, offset_x=x_count, offset_y=y_count)
                x_count += 18000000
                today_count += 1
                today_xbase += 18000000
            else:
                cn.choose_zairyou_no(msp, item, offset_x=x_count + 10000000, offset_y=y_count + 15000000)
                cn.name_sokutenn(msp, item, route_name, offset_x=x_count - 90000000, offset_y=y_count - y_count + 13500000 * y_offsetcount)
                cn.pe_hb_katauke_bend_ab(msp, offset_x=x_count, offset_y=y_count)
                first_x = x_count
                x_count += 18000000
                y_offsetcount += 1
                today_count += 1
                second_x = x_count
                flat_sum += (item.get('水平距離', 0) or 0) + (item.get('水平距離2', 0) or 0)
                haikann_sum += (item.get('配管長', 0) or 0) + (item.get('配管長2', 0) or 0)
                msp.add_line((first_x, y_count - y_count - 45000000), (first_x, y_count - y_count - 61000000))
                msp.add_line((second_x, y_count - y_count - 45000000), (second_x, y_count - y_count - 61000000))
                msp.add_line((second_x - 10000000, y_count - y_count - 45000000), (second_x - 10000000, y_count - y_count - 61000000))
                msp.add_line((first_x, y_count - y_count - 61000000), (second_x, y_count - y_count - 61000000))
                cn.flat_line1(msp, item, offset_x=x_count-13500000, offset_y=y_count - y_count - 61000000)
                cn.flat_line2(msp, item, offset_x=x_count-5000000, offset_y=y_count - y_count - 61000000)
        elif item['方向'] == 'BA':
            if today_count == 0:
                cn.pe_hb_katauke_bend_ba(msp, offset_x=x_count, offset_y=y_count)
                x_count += 18000000
                today_count += 1
                today_xbase += 18000000
            else:
                cn.choose_zairyou_no(msp, item, offset_x=x_count + 8000000, offset_y=y_count + 15000000)
                cn.name_sokutenn(msp, item, route_name, offset_x=x_count - 90000000, offset_y=y_count - y_count + 13500000 * y_offsetcount)
                cn.pe_hb_katauke_bend_ba(msp, offset_x=x_count, offset_y=y_count)
                first_x = x_count
                x_count += 18000000
                y_offsetcount += 1
                today_count += 1
                second_x = x_count
                flat_sum += (item.get('水平距離', 0) or 0) + (item.get('水平距離2', 0) or 0)
                haikann_sum += (item.get('配管長', 0) or 0) + (item.get('配管長2', 0) or 0)
                msp.add_line((first_x, y_count - y_count - 45000000), (first_x, y_count - y_count - 61000000))
                msp.add_line((second_x, y_count - y_count - 45000000), (second_x, y_count - y_count - 61000000))
                msp.add_line((second_x - 8000000, y_count - y_count - 45000000), (second_x - 8000000, y_count - y_count - 61000000))
                msp.add_line((first_x, y_count - y_count - 61000000), (second_x, y_count - y_count - 61000000))
                cn.flat_line1(msp, item, offset_x=x_count-13000000, offset_y=y_count - y_count - 61000000)
                cn.flat_line2(msp, item, offset_x=x_count-4000000, offset_y=y_count - y_count - 61000000)

    elif (
        'PE' in name
        and '挿し口付' in name
        and '丁字管' in name
        and '空気弁付消火栓' in name
    ):
        if item['方向'] == 'BB':
            if today_count == 0:
                cn.pe_ft_ryousashi_airfire_bb(msp, offset_x=x_count, offset_y=y_count)
                x_count += 16000000
                today_count += 1
                today_xbase += 16000000
            else:
                cn.choose_zairyou_no(msp, item, offset_x=x_count + 8000000, offset_y=y_count + 38000000)
                cn.name_sokutenn(msp, item, route_name, offset_x=x_count - 104000000, offset_y=y_count - y_count + 13500000 * y_offsetcount)
                cn.pe_ft_ryousashi_airfire_bb(msp, offset_x=x_count, offset_y=y_count)
                first_x = x_count
                x_count += 16000000
                y_offsetcount += 1
                today_count += 1
                second_x = x_count
                flat_sum += (item.get('水平距離', 0) or 0) + (item.get('水平距離2', 0) or 0)
                haikann_sum += (item.get('配管長', 0) or 0) + (item.get('配管長2', 0) or 0)
                msp.add_line((first_x, y_count - y_count - 45000000), (first_x, y_count - y_count - 61000000))
                msp.add_line((second_x, y_count - y_count - 45000000), (second_x, y_count - y_count - 61000000))
                msp.add_line((second_x - 8000000, y_count - y_count - 45000000), (second_x - 8000000, y_count - y_count - 61000000))
                msp.add_line((first_x, y_count - y_count - 61000000), (second_x, y_count - y_count - 61000000))
                cn.flat_line1(msp, item, offset_x=x_count-12000000, offset_y=y_count - y_count - 61000000)
                cn.flat_line2(msp, item, offset_x=x_count-4000000, offset_y=y_count - y_count - 61000000)
        elif item['方向'] == 'AA':
            if today_count == 0:
                cn.pe_ft_ryouuke_airfire_aa(msp, offset_x=x_count, offset_y=y_count)
                x_count += 20000000
                today_count += 1
                today_xbase += 20000000
            else:
                cn.choose_zairyou_no(msp, item, offset_x=x_count + 10000000, offset_y=y_count + 38000000)
                cn.name_sokutenn(msp, item, route_name, offset_x=x_count - 104000000, offset_y=y_count - y_count + 13500000 * y_offsetcount)
                cn.pe_ft_ryouuke_airfire_aa(msp, offset_x=x_count, offset_y=y_count)
                first_x = x_count
                x_count += 20000000
                y_offsetcount += 1
                today_count += 1
                second_x = x_count
                flat_sum += (item.get('水平距離', 0) or 0) + (item.get('水平距離2', 0) or 0)
                haikann_sum += (item.get('配管長', 0) or 0) + (item.get('配管長2', 0) or 0)
                msp.add_line((first_x, y_count - y_count - 45000000), (first_x, y_count - y_count - 61000000))
                msp.add_line((second_x, y_count - y_count - 45000000), (second_x, y_count - y_count - 61000000))
                msp.add_line((second_x - 10000000, y_count - y_count - 45000000), (second_x - 10000000, y_count - y_count - 61000000))
                msp.add_line((first_x, y_count - y_count - 61000000), (second_x, y_count - y_count - 61000000))
                cn.flat_line1(msp, item, offset_x=x_count-15000000, offset_y=y_count - y_count - 61000000)
                cn.flat_line2(msp, item, offset_x=x_count-5500000, offset_y=y_count - y_count - 61000000)

    elif (
        'PE' in name
        and '挿し口付' in name
        and '丁字管' in name
        and '空気弁' in name
    ):
        if item['方向'] == 'BB':
            if today_count == 0:
                cn.pe_ft_ryousashi_air_bb(msp, offset_x=x_count, offset_y=y_count)
                x_count += 16000000
                today_count += 1
                today_xbase += 16000000
            else:
                cn.choose_zairyou_no(msp, item, offset_x=x_count + 8000000, offset_y=y_count + 38000000)
                cn.name_sokutenn(msp, item, route_name, offset_x=x_count - 104000000, offset_y=y_count - y_count + 13500000 * y_offsetcount)
                cn.pe_ft_ryousashi_air_bb(msp, offset_x=x_count, offset_y=y_count)
                first_x = x_count
                x_count += 16000000
                y_offsetcount += 1
                today_count += 1
                second_x = x_count
                flat_sum += (item.get('水平距離', 0) or 0) + (item.get('水平距離2', 0) or 0)
                haikann_sum += (item.get('配管長', 0) or 0) + (item.get('配管長2', 0) or 0)
                msp.add_line((first_x, y_count - y_count - 45000000), (first_x, y_count - y_count - 61000000))
                msp.add_line((second_x, y_count - y_count - 45000000), (second_x, y_count - y_count - 61000000))
                msp.add_line((second_x - 8000000, y_count - y_count - 45000000), (second_x - 8000000, y_count - y_count - 61000000))
                msp.add_line((first_x, y_count - y_count - 61000000), (second_x, y_count - y_count - 61000000))
                cn.flat_line1(msp, item, offset_x=x_count-12000000, offset_y=y_count - y_count - 61000000)
                cn.flat_line2(msp, item, offset_x=x_count-4000000, offset_y=y_count - y_count - 61000000)
        elif item['方向'] == 'AA':
            if today_count == 0:
                cn.pe_ft_ryouuke_air_aa(msp, offset_x=x_count, offset_y=y_count)
                x_count += 20000000
                today_count += 1
                today_xbase += 20000000
            else:
                cn.choose_zairyou_no(msp, item, offset_x=x_count + 10000000, offset_y=y_count + 38000000)
                cn.name_sokutenn(msp, item, route_name, offset_x=x_count - 104000000, offset_y=y_count - y_count + 13500000 * y_offsetcount)
                cn.pe_ft_ryouuke_air_aa(msp, offset_x=x_count, offset_y=y_count)
                first_x = x_count
                x_count += 20000000
                y_offsetcount += 1
                today_count += 1
                second_x = x_count
                flat_sum += (item.get('水平距離', 0) or 0) + (item.get('水平距離2', 0) or 0)
                haikann_sum += (item.get('配管長', 0) or 0) + (item.get('配管長2', 0) or 0)
                msp.add_line((first_x, y_count - y_count - 45000000), (first_x, y_count - y_count - 61000000))
                msp.add_line((second_x, y_count - y_count - 45000000), (second_x, y_count - y_count - 61000000))
                msp.add_line((second_x - 10000000, y_count - y_count - 45000000), (second_x - 10000000, y_count - y_count - 61000000))
                msp.add_line((first_x, y_count - y_count - 61000000), (second_x, y_count - y_count - 61000000))
                cn.flat_line1(msp, item, offset_x=x_count-15000000, offset_y=y_count - y_count - 61000000)
                cn.flat_line2(msp, item, offset_x=x_count-5500000, offset_y=y_count - y_count - 61000000)

    elif (
        'PE' in name
        and '挿し口付' in name
        and '丁字管' in name
        and '消火栓' in name
    ):
        if item['方向'] == 'BB':
            if today_count == 0:
                cn.pe_ft_ryousashi_fire_bb(msp, offset_x=x_count, offset_y=y_count)
                x_count += 16000000
                today_count += 1
                today_xbase += 16000000
            else:
                cn.choose_zairyou_no(msp, item, offset_x=x_count + 8000000, offset_y=y_count + 38000000)
                cn.name_sokutenn(msp, item, route_name, offset_x=x_count - 104000000, offset_y=y_count - y_count + 13500000 * y_offsetcount)
                cn.pe_ft_ryousashi_fire_bb(msp, offset_x=x_count, offset_y=y_count)
                first_x = x_count
                x_count += 16000000
                y_offsetcount += 1
                today_count += 1
                second_x = x_count
                flat_sum += (item.get('水平距離', 0) or 0) + (item.get('水平距離2', 0) or 0)
                haikann_sum += (item.get('配管長', 0) or 0) + (item.get('配管長2', 0) or 0)
                msp.add_line((first_x, y_count - y_count - 45000000), (first_x, y_count - y_count - 61000000))
                msp.add_line((second_x, y_count - y_count - 45000000), (second_x, y_count - y_count - 61000000))
                msp.add_line((second_x - 8000000, y_count - y_count - 45000000), (second_x - 8000000, y_count - y_count - 61000000))
                msp.add_line((first_x, y_count - y_count - 61000000), (second_x, y_count - y_count - 61000000))
                cn.flat_line1(msp, item, offset_x=x_count-12000000, offset_y=y_count - y_count - 61000000)
                cn.flat_line2(msp, item, offset_x=x_count-4000000, offset_y=y_count - y_count - 61000000)

        elif item['方向'] == 'AA':
            if today_count == 0:
                cn.pe_ft_ryouuke_fire_aa(msp, offset_x=x_count, offset_y=y_count)
                x_count += 20000000
                today_count += 1
                today_xbase += 20000000
            else:
                cn.choose_zairyou_no(msp, item, offset_x=x_count + 10000000, offset_y=y_count + 38000000)
                cn.name_sokutenn(msp, item, route_name, offset_x=x_count - 104000000, offset_y=y_count - y_count + 13500000 * y_offsetcount)
                cn.pe_ft_ryouuke_fire_aa(msp, offset_x=x_count, offset_y=y_count)
                first_x = x_count
                x_count += 20000000
                y_offsetcount += 1
                today_count += 1
                second_x = x_count
                flat_sum += (item.get('水平距離', 0) or 0) + (item.get('水平距離2', 0) or 0)
                haikann_sum += (item.get('配管長', 0) or 0) + (item.get('配管長2', 0) or 0)
                msp.add_line((first_x, y_count - y_count - 45000000), (first_x, y_count - y_count - 61000000))
                msp.add_line((second_x, y_count - y_count - 45000000), (second_x, y_count - y_count - 61000000))
                msp.add_line((second_x - 10000000, y_count - y_count - 45000000), (second_x - 10000000, y_count - y_count - 61000000))
                msp.add_line((first_x, y_count - y_count - 61000000), (second_x, y_count - y_count - 61000000))
                cn.flat_line1(msp, item, offset_x=x_count-15000000, offset_y=y_count - y_count - 61000000)
                cn.flat_line2(msp, item, offset_x=x_count-5500000, offset_y=y_count - y_count - 61000000)

    elif (
        'EF' in name
        and 'チーズ' in name
        and '両受' in name
    ):
        if item['方向'] == 'SB':
            cn.choose_zairyou_no(msp, item, offset_x=x_count, offset_y=y_count+18000000)
            cn.pe_t_ryouuke_sb(msp,offset_x=x_count, offset_y=y_count)
            cn.name(msp, item, offset_x=x_count - 77000000, offset_y=y_count - y_count + 13500000 * y_offsetcount)
            first_x = x_count
            x_count += 8000000
            y_offsetcount += 1
            today_count += 1
            second_x = x_count
            flat_sum += item['水平距離']
            haikann_sum += item['配管長']
            msp.add_line((first_x, y_count - y_count - 45000000), (first_x, y_count - y_count - 61000000))
            msp.add_line((second_x, y_count - y_count - 45000000), (second_x, y_count - y_count - 61000000))
            msp.add_line((first_x, y_count - y_count - 61000000), (second_x, y_count - y_count - 61000000))
            cn.flat_line1(msp, item, offset_x=x_count - 4000000, offset_y=y_count - y_count - 61000000)
            if item['水平距離'] == item['配管長']:
                pass
            else:
                cn.haikann_line1(msp, item, offset_x=x_count-19500000, offset_y=y_count - y_count - 54500000)  
        elif item['方向'] == 'BE':
            if today_count == 0:
                cn.pe_t_ryouuke_be(msp, offset_x=x_count, offset_y=y_count)
                x_count += 8000000
                today_count += 1
                today_xbase += 8000000
            else:
                cn.choose_zairyou_no(msp, item, offset_x=x_count + 8000000, offset_y=y_count + 15000000)
                cn.name_sokutenn(msp, item, route_name, offset_x=x_count - 104000000, offset_y=y_count - y_count + 13500000 * y_offsetcount)
                cn.pe_t_ryouuke_be(msp, offset_x=x_count, offset_y=y_count)
                first_x = x_count
                x_count += 8000000
                y_offsetcount += 1
                today_count += 1
                second_x = x_count
                flat_sum += (item.get('水平距離', 0) or 0) + (item.get('水平距離2', 0) or 0)
                haikann_sum += (item.get('配管長', 0) or 0) + (item.get('配管長2', 0) or 0)
                msp.add_line((first_x, y_count - y_count - 45000000), (first_x, y_count - y_count - 61000000))
                msp.add_line((second_x, y_count - y_count - 45000000), (second_x, y_count - y_count - 61000000))
                msp.add_line((first_x, y_count - y_count - 61000000), (second_x, y_count - y_count - 61000000))
                cn.flat_line1(msp, item, offset_x=x_count - 4000000, offset_y=y_count - y_count - 61000000)
        elif item['方向'] == 'AA':
            if today_count == 0:
                cn.pe_t_ryouuke_aa(msp, offset_x=x_count, offset_y=y_count)
                x_count += 20000000
                today_count += 1
                today_xbase += 20000000
            else:
                cn.choose_zairyou_no(msp, item, offset_x=x_count + 10000000, offset_y=y_count + 15000000)
                cn.name_sokutenn(msp, item, route_name, offset_x=x_count - 104000000, offset_y=y_count - y_count + 13500000 * y_offsetcount)
                cn.pe_t_ryouuke_aa(msp, offset_x=x_count, offset_y=y_count)
                first_x = x_count
                x_count += 20000000
                y_offsetcount += 1
                today_count += 1
                second_x = x_count
                flat_sum += (item.get('水平距離', 0) or 0) + (item.get('水平距離2', 0) or 0)
                haikann_sum += (item.get('配管長', 0) or 0) + (item.get('配管長2', 0) or 0)
                msp.add_line((first_x, y_count - y_count - 45000000), (first_x, y_count - y_count - 61000000))
                msp.add_line((second_x, y_count - y_count - 45000000), (second_x, y_count - y_count - 61000000))
                msp.add_line((second_x - 10000000, y_count - y_count - 45000000), (second_x - 10000000, y_count - y_count - 61000000))
                msp.add_line((first_x, y_count - y_count - 61000000), (second_x, y_count - y_count - 61000000))
                cn.flat_line1(msp, item, offset_x=x_count-15000000, offset_y=y_count - y_count - 61000000)
                cn.flat_line2(msp, item, offset_x=x_count-5000000, offset_y=y_count - y_count - 61000000)
        elif item['方向'] == 'AB':
            if today_count == 0:
                cn.pe_hb_katauke_bend_ab(msp, offset_x=x_count, offset_y=y_count)
                x_count += 18000000
                today_count += 1
                today_xbase += 18000000
            else:
                cn.choose_zairyou_no(msp, item, offset_x=x_count + 10000000, offset_y=y_count + 15000000)
                cn.name_sokutenn(msp, item, route_name, offset_x=x_count - 90000000, offset_y=y_count - y_count + 13500000 * y_offsetcount)
                cn.pe_hb_katauke_bend_ab(msp, offset_x=x_count, offset_y=y_count)
                first_x = x_count
                x_count += 18000000
                y_offsetcount += 1
                today_count += 1
                second_x = x_count
                flat_sum += (item.get('水平距離', 0) or 0) + (item.get('水平距離2', 0) or 0)
                haikann_sum += (item.get('配管長', 0) or 0) + (item.get('配管長2', 0) or 0)
                msp.add_line((first_x, y_count - y_count - 45000000), (first_x, y_count - y_count - 61000000))
                msp.add_line((second_x, y_count - y_count - 45000000), (second_x, y_count - y_count - 61000000))
                msp.add_line((second_x - 10000000, y_count - y_count - 45000000), (second_x - 10000000, y_count - y_count - 61000000))
                msp.add_line((first_x, y_count - y_count - 61000000), (second_x, y_count - y_count - 61000000))
                cn.flat_line1(msp, item, offset_x=x_count-13500000, offset_y=y_count - y_count - 61000000)
                cn.flat_line2(msp, item, offset_x=x_count-2500000, offset_y=y_count - y_count - 61000000)
        elif item['方向'] == 'BA':
            if today_count == 0:
                cn.pe_hb_katauke_bend_ba(msp, offset_x=x_count, offset_y=y_count)
                x_count += 18000000
                today_count += 1
                today_xbase += 18000000
            else:
                cn.choose_zairyou_no(msp, item, offset_x=x_count + 8000000, offset_y=y_count + 15000000)
                cn.name_sokutenn(msp, item, route_name, offset_x=x_count - 90000000, offset_y=y_count - y_count + 13500000 * y_offsetcount)
                cn.pe_hb_katauke_bend_ba(msp, offset_x=x_count, offset_y=y_count)
                first_x = x_count
                x_count += 18000000
                y_offsetcount += 1
                today_count += 1
                second_x = x_count
                flat_sum += (item.get('水平距離', 0) or 0) + (item.get('水平距離2', 0) or 0)
                haikann_sum += (item.get('配管長', 0) or 0) + (item.get('配管長2', 0) or 0)
                msp.add_line((first_x, y_count - y_count - 45000000), (first_x, y_count - y_count - 61000000))
                msp.add_line((second_x, y_count - y_count - 45000000), (second_x, y_count - y_count - 61000000))
                msp.add_line((second_x - 8000000, y_count - y_count - 45000000), (second_x - 8000000, y_count - y_count - 61000000))
                msp.add_line((first_x, y_count - y_count - 61000000), (second_x, y_count - y_count - 61000000))
                cn.flat_line1(msp, item, offset_x=x_count-13000000, offset_y=y_count - y_count - 61000000)
                cn.flat_line2(msp, item, offset_x=x_count-2500000, offset_y=y_count - y_count - 61000000)

    elif (
        'EF' in name
        and '両挿仕切弁' in name
    ):
        if item['方向'] == 'BB':
            if today_count == 0:
                cn.pe_ryousashi_ben_BB(msp, offset_x=x_count, offset_y=y_count)
                x_count += 25000000
                today_count += 1
                today_xbase += 25000000
            else:
                cn.choose_zairyou_no(msp, item, offset_x=x_count + 12500000, offset_y=y_count + 15000000)
                cn.name_sokutenn(msp, item, route_name, offset_x=x_count - 90000000, offset_y=y_count - y_count + 13500000 * y_offsetcount)
                cn.pe_ryousashi_ben_BB(msp, offset_x=x_count, offset_y=y_count)
                first_x = x_count
                x_count += 25000000
                y_offsetcount += 1
                today_count += 1
                second_x = x_count
                flat_sum += (item.get('水平距離', 0) or 0) + (item.get('水平距離2', 0) or 0)
                haikann_sum += (item.get('配管長', 0) or 0) + (item.get('配管長2', 0) or 0)
                msp.add_line((first_x, y_count - y_count - 45000000), (first_x, y_count - y_count - 61000000))
                msp.add_line((second_x, y_count - y_count - 45000000), (second_x, y_count - y_count - 61000000))
                msp.add_line((second_x - 12500000, y_count - y_count - 45000000), (second_x - 12500000, y_count - y_count - 61000000))
                msp.add_line((first_x, y_count - y_count - 61000000), (second_x, y_count - y_count - 61000000))
                cn.flat_line1(msp, item, offset_x=x_count-18750000, offset_y=y_count - y_count - 61000000)
                cn.flat_line2(msp, item, offset_x=x_count-6250000, offset_y=y_count - y_count - 61000000)

    elif (
        '仕切弁' in name
    ):
        if item['方向'] == 'FF':
            if today_count == 0:
                cn.soft_ben_FF(msp, offset_x=x_count, offset_y=y_count)
                x_count += 19000000
                today_count += 1
                today_xbase += 19000000
            else:
                cn.choose_zairyou_no(msp, item, offset_x=x_count + 9000000, offset_y=y_count + 15000000)
                cn.name_sokutenn(msp, item, route_name, offset_x=x_count - 90000000, offset_y=y_count - y_count + 13500000 * y_offsetcount)
                cn.soft_ben_FF(msp, offset_x=x_count, offset_y=y_count)
                first_x = x_count
                x_count += 19000000
                y_offsetcount += 1
                today_count += 1
                second_x = x_count
                flat_sum += (item.get('水平距離', 0) or 0) + (item.get('水平距離2', 0) or 0)
                haikann_sum += (item.get('配管長', 0) or 0) + (item.get('配管長2', 0) or 0)
                msp.add_line((first_x, y_count - y_count - 45000000), (first_x, y_count - y_count - 61000000))
                msp.add_line((second_x, y_count - y_count - 45000000), (second_x, y_count - y_count - 61000000))
                msp.add_line((first_x, y_count - y_count - 61000000), (second_x, y_count - y_count - 61000000))
                cn.flat_line1(msp, item, offset_x=x_count-9000000, offset_y=y_count - y_count - 61000000)

    elif (
        'EF' in name
        and 'フランジ短管' in name
    ):
        if item['方向'] == 'SB':
            cn.choose_zairyou_no(msp, item, offset_x=x_count + 5000000, offset_y=y_count+15000000)
            cn.pe_franjitann_FB(msp, offset_x=x_count, offset_y=y_count)
            cn.name(msp, item, offset_x=x_count - 55000000, offset_y=y_count + 8000000 + 13500000 * y_offsetcount)
            first_x = x_count
            x_count += 10000000
            y_offsetcount += 1
            today_count += 1
            second_x = x_count
            flat_sum += (item.get('水平距離', 0) or 0) + (item.get('水平距離2', 0) or 0)
            haikann_sum += (item.get('配管長', 0) or 0) + (item.get('配管長2', 0) or 0)
            msp.add_line((first_x, y_count - y_count - 45000000), (first_x, y_count - y_count - 61000000))
            msp.add_line((second_x, y_count - y_count - 45000000), (second_x, y_count - y_count - 61000000))
            msp.add_line((first_x, y_count - y_count - 61000000), (second_x, y_count - y_count - 61000000))
            cn.flat_line1(msp, item, offset_x=x_count - 5000000, offset_y=y_count - y_count - 61000000)

        elif item['方向'] == 'FB':
            if today_count == 0:
                cn.pe_franjitann_FB(msp, offset_x=x_count, offset_y=y_count)
                x_count += 10000000
                today_count += 1
                today_xbase += 10000000
            else:
                cn.choose_zairyou_no(msp, item, offset_x=x_count + 5000000, offset_y=y_count + 15000000)
                cn.name(msp, item, offset_x=x_count - 90000000, offset_y=y_count - y_count + 13500000 * y_offsetcount)
                cn.pe_franjitann_FB(msp, offset_x=x_count, offset_y=y_count)
                first_x = x_count
                x_count += 10000000
                y_offsetcount += 1
                today_count += 1
                second_x = x_count
                flat_sum += (item.get('水平距離', 0) or 0) + (item.get('水平距離2', 0) or 0)
                haikann_sum += (item.get('配管長', 0) or 0) + (item.get('配管長2', 0) or 0)
                msp.add_line((first_x, y_count - y_count - 45000000), (first_x, y_count - y_count - 61000000))
                msp.add_line((second_x, y_count - y_count - 45000000), (second_x, y_count - y_count - 61000000))
                msp.add_line((first_x, y_count - y_count - 61000000), (second_x, y_count - y_count - 61000000))
                cn.flat_line1(msp, item, offset_x=x_count - 5000000, offset_y=y_count - y_count - 61000000)

        elif item['方向'] == 'BF':
            if today_count == 0:
                cn.pe_franjitann_BF(msp, offset_x=x_count, offset_y=y_count)
                x_count += 11000000
                today_count += 1
                today_xbase += 11000000
            else:
                cn.choose_zairyou_no(msp, item, offset_x=x_count + 5000000, offset_y=y_count + 15000000)
                cn.name(msp, item, offset_x=x_count - 90000000, offset_y=y_count - y_count + 13500000 * y_offsetcount)
                cn.pe_franjitann_BF(msp, offset_x=x_count, offset_y=y_count)
                first_x = x_count
                x_count += 11000000
                y_offsetcount += 1
                today_count += 1
                second_x = x_count
                flat_sum += (item.get('水平距離', 0) or 0) + (item.get('水平距離2', 0) or 0)
                haikann_sum += (item.get('配管長', 0) or 0) + (item.get('配管長2', 0) or 0)
                msp.add_line((first_x, y_count - y_count - 45000000), (first_x, y_count - y_count - 61000000))
                msp.add_line((second_x, y_count - y_count - 45000000), (second_x, y_count - y_count - 61000000))
                msp.add_line((first_x, y_count - y_count - 61000000), (second_x, y_count - y_count - 61000000))
                cn.flat_line1(msp, item, offset_x=x_count - 5000000, offset_y=y_count - y_count - 61000000)

    elif (
        'EF' in name
        and 'フランジ' in name
    ):
        if item['方向'] == 'FA':
            if today_count == 0:
                cn.pe_franji_FA(msp, offset_x=x_count, offset_y=y_count)
                x_count += 15000000
                today_count += 1
                today_xbase += 15000000
            else:
                cn.choose_zairyou_no(msp, item, offset_x=x_count + 7500000, offset_y=y_count + 15000000)
                cn.name(msp, item, offset_x=x_count - 90000000, offset_y=y_count - y_count + 13500000 * y_offsetcount)
                cn.pe_franji_FA(msp, offset_x=x_count, offset_y=y_count)
                first_x = x_count
                x_count += 15000000
                y_offsetcount += 1
                today_count += 1
                second_x = x_count
                flat_sum += (item.get('水平距離', 0) or 0) + (item.get('水平距離2', 0) or 0)
                haikann_sum += (item.get('配管長', 0) or 0) + (item.get('配管長2', 0) or 0)
                msp.add_line((first_x, y_count - y_count - 45000000), (first_x, y_count - y_count - 61000000))
                msp.add_line((second_x, y_count - y_count - 45000000), (second_x, y_count - y_count - 61000000))
                msp.add_line((first_x, y_count - y_count - 61000000), (second_x, y_count - y_count - 61000000))
                cn.flat_line1(msp, item, offset_x=x_count - 7500000, offset_y=y_count - y_count - 61000000)

        elif item['方向'] == 'AF':
            if today_count == 0:
                cn.pe_franji_AF(msp, offset_x=x_count, offset_y=y_count)
                x_count += 16000000
                today_count += 1
                today_xbase += 16000000
            else:
                cn.choose_zairyou_no(msp, item, offset_x=x_count + 7500000, offset_y=y_count + 15000000)
                cn.name(msp, item, offset_x=x_count - 90000000, offset_y=y_count - y_count + 13500000 * y_offsetcount)
                cn.pe_franji_AF(msp, offset_x=x_count, offset_y=y_count)
                first_x = x_count
                x_count += 16000000
                y_offsetcount += 1
                today_count += 1
                second_x = x_count
                flat_sum += (item.get('水平距離', 0) or 0) + (item.get('水平距離2', 0) or 0)
                haikann_sum += (item.get('配管長', 0) or 0) + (item.get('配管長2', 0) or 0)
                msp.add_line((first_x, y_count - y_count - 45000000), (first_x, y_count - y_count - 61000000))
                msp.add_line((second_x, y_count - y_count - 45000000), (second_x, y_count - y_count - 61000000))
                msp.add_line((first_x, y_count - y_count - 61000000), (second_x, y_count - y_count - 61000000))
                cn.flat_line1(msp, item, offset_x=x_count - 7500000, offset_y=y_count - y_count - 61000000)

    elif (
        'EF' in name
        and 'レデューサー' in name
        and '小径化' in name
    ):
        if item['方向'] == 'BB':
            if today_count == 0:
                cn.pe_reducer_updown_bb(msp, offset_x=x_count, offset_y=y_count)
                x_count += 15000000
                today_count += 1
                today_xbase += 15000000
            else:
                cn.choose_zairyou_no(msp, item, offset_x=x_count + 7500000, offset_y=y_count + 15000000)
                cn.name(msp, item, offset_x=x_count - 90000000, offset_y=y_count - y_count + 13500000 * y_offsetcount)
                cn.pe_reducer_updown_bb(msp, offset_x=x_count, offset_y=y_count)
                first_x = x_count
                x_count += 15000000
                y_offsetcount += 1
                today_count += 1
                second_x = x_count
                flat_sum += (item.get('水平距離', 0) or 0) + (item.get('水平距離2', 0) or 0)
                haikann_sum += (item.get('配管長', 0) or 0) + (item.get('配管長2', 0) or 0)
                msp.add_line((first_x, y_count - y_count - 45000000), (first_x, y_count - y_count - 61000000))
                msp.add_line((second_x, y_count - y_count - 45000000), (second_x, y_count - y_count - 61000000))
                msp.add_line((first_x, y_count - y_count - 61000000), (second_x, y_count - y_count - 61000000))
                flat_line1_value = item.get('水平距離', 0) or 0
                flat_line2_value = item.get('水平距離2', 0) or 0
                goukei = flat_line1_value + flat_line2_value
                msp.add_text(str(goukei), dxfattribs={"height": 6000000}).set_placement(
                    (x_count - 12200000, y_count - y_count - 61000000)
                )

        elif item['方向'] == 'AF':
            if today_count == 0:
                cn.pe_franji_AF(msp, offset_x=x_count, offset_y=y_count)
                x_count += 15000000
                today_count += 1
                today_xbase += 15000000
            else:
                cn.choose_zairyou_no(msp, item, offset_x=x_count + 7500000, offset_y=y_count + 15000000)
                cn.name(msp, item, offset_x=x_count - 90000000, offset_y=y_count - y_count + 13500000 * y_offsetcount)
                cn.pe_franji_AF(msp, offset_x=x_count, offset_y=y_count)
                first_x = x_count
                x_count += 15000000
                y_offsetcount += 1
                today_count += 1
                second_x = x_count
                flat_sum += (item.get('水平距離', 0) or 0) + (item.get('水平距離2', 0) or 0)
                haikann_sum += (item.get('配管長', 0) or 0) + (item.get('配管長2', 0) or 0)
                msp.add_line((first_x, y_count - y_count - 45000000), (first_x, y_count - y_count - 61000000))
                msp.add_line((second_x, y_count - y_count - 45000000), (second_x, y_count - y_count - 61000000))
                msp.add_line((first_x, y_count - y_count - 61000000), (second_x, y_count - y_count - 61000000))
                cn.flat_line1(msp, item, offset_x=x_count - 7500000, offset_y=y_count - y_count - 61000000)

    elif (
        'EF' in name
        and 'レデューサー' in name
        and '大径化' in name
    ):
        if item['方向'] == 'BB':
            if today_count == 0:
                cn.pe_reducer_downup_bb(msp, offset_x=x_count, offset_y=y_count)
                x_count += 15000000
                today_count += 1
                today_xbase += 15000000
            else:
                cn.choose_zairyou_no(msp, item, offset_x=x_count + 7500000, offset_y=y_count + 15000000)
                cn.name(msp, item, offset_x=x_count - 90000000, offset_y=y_count - y_count + 13500000 * y_offsetcount)
                cn.pe_reducer_downup_bb(msp, offset_x=x_count, offset_y=y_count)
                first_x = x_count
                x_count += 15000000
                y_offsetcount += 1
                today_count += 1
                second_x = x_count
                flat_sum += (item.get('水平距離', 0) or 0) + (item.get('水平距離2', 0) or 0)
                haikann_sum += (item.get('配管長', 0) or 0) + (item.get('配管長2', 0) or 0)
                msp.add_line((first_x, y_count - y_count - 45000000), (first_x, y_count - y_count - 61000000))
                msp.add_line((second_x, y_count - y_count - 45000000), (second_x, y_count - y_count - 61000000))
                msp.add_line((first_x, y_count - y_count - 61000000), (second_x, y_count - y_count - 61000000))
                flat_line1_value = item.get('水平距離', 0) or 0
                flat_line2_value = item.get('水平距離2', 0) or 0
                goukei = flat_line1_value + flat_line2_value
                msp.add_text(str(goukei), dxfattribs={"height": 6000000}).set_placement(
                    (x_count - 12200000, y_count - y_count - 61000000)
                )

        elif item['方向'] == 'AF':
            if today_count == 0:
                cn.pe_franji_AF(msp, offset_x=x_count, offset_y=y_count)
                x_count += 15000000
                today_count += 1
                today_xbase += 15000000
            else:
                cn.choose_zairyou_no(msp, item, offset_x=x_count + 7500000, offset_y=y_count + 15000000)
                cn.name(msp, item, offset_x=x_count - 90000000, offset_y=y_count - y_count + 13500000 * y_offsetcount)
                cn.pe_franji_AF(msp, offset_x=x_count, offset_y=y_count)
                first_x = x_count
                x_count += 15000000
                y_offsetcount += 1
                today_count += 1
                second_x = x_count
                flat_sum += (item.get('水平距離', 0) or 0) + (item.get('水平距離2', 0) or 0)
                haikann_sum += (item.get('配管長', 0) or 0) + (item.get('配管長2', 0) or 0)
                msp.add_line((first_x, y_count - y_count - 45000000), (first_x, y_count - y_count - 61000000))
                msp.add_line((second_x, y_count - y_count - 45000000), (second_x, y_count - y_count - 61000000))
                msp.add_line((first_x, y_count - y_count - 61000000), (second_x, y_count - y_count - 61000000))
                cn.flat_line1(msp, item, offset_x=x_count - 7500000, offset_y=y_count - y_count - 61000000)

    elif (
        'EF' in name
        and 'ソケット' in name
        and '(V)' in name
        and down_90
    ):
        if item['方向'] == 'AA':
            if today_count == 0:
                cn.pe_vb_socket_90_down_aa(msp, offset_x=x_count, offset_y=y_count + 5000000)
                today_count += 1         
            else:
                cn.choose_zairyou_no(msp, item, offset_x=x_count, offset_y=y_count + 15000000)
                cn.name(msp, item, offset_x=x_count - 70000000, offset_y=y_count - y_count + 13500000 * y_offsetcount)
                cn.pe_vb_socket_90_down_aa(msp, offset_x=x_count, offset_y=y_count + 5000000)
                y_offsetcount += 1
                today_count += 1

    elif (
        'EF' in name
        and 'ソケット' in name
        and '(V)' in name
        and up_90
    ):
        if item['方向'] == 'AA':
            if today_count == 0:
                cn.pe_vb_socket_90_up_aa(msp, offset_x=x_count, offset_y=y_count - 5000000)
                today_count += 1         
            else:
                cn.choose_zairyou_no(msp, item, offset_x=x_count, offset_y=y_count + 15000000)
                cn.name(msp, item, offset_x=x_count - 70000000, offset_y=y_count - y_count + 13500000 * y_offsetcount)
                cn.pe_vb_socket_90_up_aa(msp, offset_x=x_count, offset_y=y_count - 5000000)
                y_offsetcount += 1
                today_count += 1

    elif (
        'EF' in name
        and 'ソケット' in name
        and '(V)' in name
        and down_45
    ):
        if item['方向'] == 'AA':
            if today_count == 0:
                cn.pe_vb_socket_45_down_aa(msp, offset_x=x_count - 3535000, offset_y=y_count + 3535000)
                today_count += 1         
            else:
                cn.choose_zairyou_no(msp, item, offset_x=x_count, offset_y=y_count + 15000000)
                cn.name(msp, item, offset_x=x_count - 70000000, offset_y=y_count - y_count + 13500000 * y_offsetcount)
                cn.pe_vb_socket_45_down_aa(msp, offset_x=x_count - 3535000, offset_y=y_count + 3535000)
                y_offsetcount += 1
                today_count += 1

    elif (
        'EF' in name
        and 'ソケット' in name
        and '(V)' in name
        and up_45
    ):
        if item['方向'] == 'AA':
            if today_count == 0:
                cn.pe_vb_socket_45_up_aa(msp, offset_x=x_count - 3535000, offset_y=y_count - 3535000)
                today_count += 1
            else:
                cn.choose_zairyou_no(msp, item, offset_x=x_count, offset_y=y_count + 15000000)
                cn.name(msp, item, offset_x=x_count - 70000000, offset_y=y_count - y_count + 13500000 * y_offsetcount)
                cn.pe_vb_socket_45_up_aa(msp, offset_x=x_count - 3535000, offset_y=y_count - 3535000)
                y_offsetcount += 1
                today_count += 1

    elif (
        'EF' in name
        and 'ソケット' in name
        and '(V)' in name
        and down_22
    ):
        if item['方向'] == 'AA':
            if today_count == 0:
                cn.pe_vb_socket_22_down_aa(msp, offset_x=x_count - 4620000, offset_y=y_count + 1915000)
                today_count += 1
            else:
                cn.choose_zairyou_no(msp, item, offset_x=x_count, offset_y=y_count + 15000000)
                cn.name(msp, item, offset_x=x_count - 70000000, offset_y=y_count - y_count + 13500000 * y_offsetcount)
                cn.pe_vb_socket_22_down_aa(msp, offset_x=x_count - 4620000, offset_y=y_count + 1915000)
                y_offsetcount += 1
                today_count += 1

    elif (
        'EF' in name
        and 'ソケット' in name
        and '(V)' in name
        and up_22
    ):
        if item['方向'] == 'AA':
            if today_count == 0:
                cn.pe_vb_socket_22_up_aa(msp, offset_x=x_count - 4620000, offset_y=y_count - 1915000)
                today_count += 1
            else:
                cn.choose_zairyou_no(msp, item, offset_x=x_count, offset_y=y_count + 15000000)
                cn.name(msp, item, offset_x=x_count - 70000000, offset_y=y_count - y_count + 13500000 * y_offsetcount)
                cn.pe_vb_socket_22_up_aa(msp, offset_x=x_count - 4620000, offset_y=y_count - 1915000)
                y_offsetcount += 1
                today_count += 1

    elif (
        'EF' in name
        and 'ソケット' in name
        and '(V)' in name
        and down_11
    ):
        if item['方向'] == 'AA':
            if today_count == 0:
                cn.pe_vb_socket_11_down_aa(msp, offset_x=x_count - 4905000, offset_y=y_count + 975500)
                today_count += 1
            else:
                cn.choose_zairyou_no(msp, item, offset_x=x_count, offset_y=y_count + 15000000)
                cn.name(msp, item, offset_x=x_count - 70000000, offset_y=y_count - y_count + 13500000 * y_offsetcount)
                cn.pe_vb_socket_11_down_aa(msp, offset_x=x_count - 4905000, offset_y=y_count + 975500)
                y_offsetcount += 1
                today_count += 1

    elif (
        'EF' in name
        and 'ソケット' in name
        and '(V)' in name
        and up_11
    ):
        if item['方向'] == 'AA':
            if today_count == 0:
                cn.pe_vb_socket_11_up_aa(msp, offset_x=x_count - 4905000, offset_y=y_count - 975500)
                today_count += 1
            else:
                cn.choose_zairyou_no(msp, item, offset_x=x_count, offset_y=y_count + 15000000)
                cn.name(msp, item, offset_x=x_count - 70000000, offset_y=y_count - y_count + 13500000 * y_offsetcount)
                cn.pe_vb_socket_11_up_aa(msp, offset_x=x_count - 4905000, offset_y=y_count - 975500)
                y_offsetcount += 1
                today_count += 1

    elif (
        'EF' in name
        and 'ソケット' in name
    ):
        if item['方向'] == 'AA':
            if today_count == 0:
                cn.pe_socket_aa(msp, offset_x=x_count -4000000, offset_y=y_count)
                today_count += 1         
            else:
                cn.choose_zairyou_no(msp, item, offset_x=x_count, offset_y=y_count + 15000000)
                cn.name(msp, item, offset_x=x_count - 70000000, offset_y=y_count - y_count + 13500000 * y_offsetcount)
                cn.pe_socket_aa(msp, offset_x=x_count - 4000000, offset_y=y_count)
                y_offsetcount += 1
                today_count += 1

    elif (
        'メカニカルソケット' in name
    ):
        if item['方向'] == 'SA':
            cn.choose_zairyou_no(msp, item, offset_x=x_count + 5000000, offset_y=y_count + 15000000)
            cn.name(msp, item, offset_x=x_count - 70000000, offset_y=y_count)
            cn.meka_socket_aa(msp, offset_x=x_count, offset_y=y_count)
            first_x = x_count
            y_offsetcount += 1
            today_count += 1
            x_count += 5000000
            today_xbase += 5000000
        elif item['方向'] == 'AA':
            if today_count == 0:
                cn.meka_socket_aa(msp, offset_x=x_count, offset_y=y_count)
                x_count += 5000000
                today_count += 1
                today_xbase += 5000000
            else:      
                cn.choose_zairyou_no(msp, item, offset_x=x_count + 5000000, offset_y=y_count + 15000000)
                cn.name(msp, item,offset_x=x_count - 100000000, offset_y=y_count - y_count + 13500000 * y_offsetcount)
                cn.meka_socket_aa(msp, offset_x=x_count - 5000000, offset_y=y_count)
                y_offsetcount += 1
                today_count += 1
        elif item['方向'] == 'AE':
            if today_count == 0:
                cn.meka_socket_aa(msp, offset_x=x_count, offset_y=y_count)
                today_count += 1
            else:      
                cn.choose_zairyou_no(msp, item, offset_x=x_count + 5000000, offset_y=y_count + 15000000)
                cn.name(msp, item,offset_x=x_count - 100000000, offset_y=y_count - y_count + 13500000 * y_offsetcount)
                cn.meka_socket_aa(msp, offset_x=x_count - 5000000, offset_y=y_count)
                y_offsetcount += 1
                today_count += 1

    elif (
        'メカニカルキャップ' in name
    ):
        if item['方向'] == 'SA':
            cn.choose_zairyou_no(msp, item, offset_x=x_count + 2500000, offset_y=y_count + 15000000)
            cn.name(msp, item, offset_x=x_count - 70000000, offset_y=y_count)
            cn.meka_cap_ae(msp, offset_x=x_count, offset_y=y_count)
            first_x = x_count
            y_offsetcount += 1
            today_count += 1
            x_count += 2500000
            today_xbase += 2500000
        elif item['方向'] == 'AA':
            if today_count == 0:
                cn.meka_cap_ae(msp, offset_x=x_count, offset_y=y_count)
                x_count += 2500000
                today_count += 1
                today_xbase += 2500000
            else:      
                cn.choose_zairyou_no(msp, item, offset_x=x_count + 2500000, offset_y=y_count + 15000000)
                cn.name(msp, item,offset_x=x_count - 100000000, offset_y=y_count - y_count + 13500000 * y_offsetcount)
                cn.meka_cap_ae(msp, offset_x=x_count - 5000000, offset_y=y_count)
                y_offsetcount += 1
                today_count += 1
        elif item['方向'] == 'AE':
            if today_count == 0:
                cn.meka_cap_ae(msp, offset_x=x_count, offset_y=y_count)
                today_count += 1
            else:
                cn.choose_zairyou_no(msp, item, offset_x=x_count + 2500000, offset_y=y_count + 15000000)
                cn.name(msp, item,offset_x=x_count - 100000000, offset_y=y_count - y_count + 13500000 * y_offsetcount)
                cn.meka_cap_ae(msp, offset_x=x_count - 5000000, offset_y=y_count)
                y_offsetcount += 1
                today_count += 1

#全体の寸法を表示
msp.add_line((today_xbase - 160000000 , y_count - y_count - 61000000), (today_xbase - 160000000, y_count - y_count - 77000000))
msp.add_line((x_count, y_count - y_count - 61000000), (x_count, y_count - y_count - 77000000))
msp.add_line((today_xbase - 160000000, y_count - y_count - 77000000), (x_count, y_count - y_count - 77000000))
msp.add_text(f"{flat_sum}", dxfattribs={
'insert': ((x_count + today_xbase - 160000000)/2 - 2600000, y_count - y_count - 77000000),
'height': 6000000,
})
if haikann_sum > flat_sum:
    msp.add_text(f"({haikann_sum})", dxfattribs={
    'insert': ((x_count + today_xbase - 160000000)/2 - 5500000, y_count - y_count - 70400000),
    'height': 6000000,
    })
else:
    pass

#終点のナンバー表示
if filtered_data[-1]['方向'] == 'AE' or filtered_data[-1]['方向'] == 'BE' :
    pass
else:
    msp.add_line((x_count, y_count), (x_count + 15000000, y_count + 13000000))
    msp.add_line((x_count + 15000000, y_count + 13000000), (x_count + 65000000, y_count + 13000000))
    if filtered_data[-1]['測点2'] == None :
        msp.add_text(f"{route_name} {filtered_data[-1]['測点']}", dxfattribs={
        'insert': (x_count + 15000000, y_count + 13000000),
        'height': 6000000,
        })
        msp.add_text("後日施工", dxfattribs={
        'insert': (x_count + 15000000, y_count + 5000000),
        'height': 6000000,
        })
    else:
        msp.add_text(f"{route_name} {filtered_data[-1]['測点2']}", dxfattribs={
        'insert': (x_count + 15000000, y_count + 13000000),
        'height': 6000000,
        })
        msp.add_text("後日施工", dxfattribs={
        'insert': (x_count + 15000000, y_count + 5000000),
        'height': 6000000,
        })


# base_x = 0 #+ offset_x
# base_y = 0 #+ offset_y

# # ライナ文字を入れる（中央揃え）
# center_pos = (base_x, base_y)

# text = msp.add_text(
#     "ﾗｲﾅ",
#     dxfattribs={
#         "insert": center_pos,   # いちおう insert も同じ座標にしておく
#         "height": 6000000,
#     }
# )

# # TEXTエンティティのアライメント設定（DXF標準の属性）
# text.dxf.halign = 1          # 0=左, 1=中央, 2=右, ...
# text.dxf.valign = 2          # 0=ベースライン, 1=下, 2=中央, 3=上
# text.dxf.align_point = center_pos  # ここが「中央になる座標」

# msp.add_line((base_x, base_y), (base_x + 75000000, base_y))

# 保存先フォルダとファイル名を指定
output_dir = r"C:\\JWW\\配管長略図"
os.makedirs(output_dir, exist_ok=True)  # outputフォルダが存在しない場合は作成
output_file_path = os.path.join(output_dir, "text_sample.dxf")

# 保存
doc.saveas(output_file_path, encoding='cp932')
print(f"DXFファイルを保存しました: {output_file_path}")