# =====================================================
# 地方競馬(NAR)予想 v137_001  ── オッズ×統計 アンサンブル版(学習モデル不使用)
#
# ★v137_001: v136_021 から予想の中身を全面的に作り替えた別バージョン。
#   出力ファイル(CSV / カラフルExcel / 研究用HTML・PDF / note用HTML・PDF / 統計観察記録CSV)と
#   予想表のスタイルは v136_021 と同じ。
#
#   【入力】(学習モデル keiba_prob_model / 予想CSV / horselist_features.py は使わない)
#     1) 単勝オッズファイル  nar_shutuba_all_YYYYMMDD.csv(netkeiba 出馬表: 枠/馬番/馬名/性齢/斤量/騎手/
#        厩舎/オッズ/人気/レースID)。GUI『オッズCSVを開く』・一括予想はこのファイルを読む。
#        レースID(YYYY+場コード+MMDD+RR)から 日付・競馬場・R を決める。
#     2) ホースリスト  *_horselist.csv(オッズCSVと同じフォルダ)。成績・最高タイム・騎手・減量記号。
#        (競馬場, 日付, R, 馬番)で結合し、馬名が違う行は結合しない。
#     3) 出馬表HTML  R01_場_日付.html(GUI『出馬表HTMLフォルダを読込』。オッズCSVのフォルダに
#        R??_*.html があれば自動で読む)。bado_stat_model.py/.json で『統計勝率』を計算し、
#        発走時刻・距離・出走取消もここから取る。
#
#   【単勝確率】= アンサンブル確率
#     ・オッズ勝率 = 補正前の単勝オッズ(1/オッズ)をレース内で正規化(ODDS_PROB_POWER で人気薄の割引も可)
#     ・統計勝率   = 出馬表HTMLの統計モデル(bado_stat_model)
#     ・単勝確率   = オッズ勝率^W_ODDS × 統計勝率^W_STAT をレース内で正規化(対数線形の融合)
#       W_ODDS/W_STAT の既定は 0.805/0.345(統計比率0.3・合計1.15。2026/09/17〜29の506Rで的中率と回収率のバランスが最良)。
#       統計勝率が無いレースはオッズ勝率^1.10(ENS_ODDS_ONLY_POWER。注意書きを出す)。
#       単勝確率を使う基準値(軸級・信頼度・投資/有力/見送り・単指数90・本線条件・穴・色分け)は、
#       尖らせ前と同じレースが選ばれるよう再調整済み(各定数のコメントに旧値)。
#     ・複勝確率/連対率/馬連・ワイドの的中率 = 単勝確率から Harville 型で計算。2着・3着は
#       確率を λ2=0.81/λ3=0.65 乗して割り引く(人気馬が2・3着に残る確率の過大評価を抑える補正)。
#
#   【予想表の単オッズ】= 予想オッズ(アンサンブル)
#     補正前の単勝オッズ と 人気オッズ(人気になりそうな項目から算出)を対数で加重平均し
#     (POP_ODDS_W_RAW)、補正前オッズと同じ控除率に揃えたもの。人気=この予想オッズの順。
#     人気オッズ = 騎手の評価(当日の他の騎乗馬の人気・騎乗数)/減量騎手/全成績の勝率/
#       当地の複勝率/騎手×馬のコンビ成績/キャリア/最高タイム から人気の集まり方を推定した単勝オッズ。
#       係数は 2026/09/29 の NAR 5場 59R の市場オッズに当てはめた値(場を1つずつ抜いた検証で
#       順位相関 0.47〜0.73)。払戻率は単勝80%。
#
#   【騎手指数】(新ロジック・0〜100、50=その場の平均、60以上=好騎手)
#     当日同じ場で乗る『他の騎乗馬』の人気(オッズ勝率×頭数の対数平均。本馬は除く=自分の人気を
#     混ぜない)・騎乗数・減量騎手・この馬とのコンビ成績を、人気オッズの係数で合成し偏差値化。
#   【紐馬指数】(新ロジック・0〜100、レース内で最も紐向きの馬=100)
#     紐は『◎が勝ったときに2・3着に来る馬』なので、(複勝確率−単勝確率)=2〜3着に来る確率を土台に、
#     配当の大きさとして予想オッズの √ を掛ける(来やすさと配当のバランス)。
#   【単指数】(100点満点・90以上=鉄板クラス) = 単勝確率を 100*(1-10^(-p/0.625)) で換算(単勝確率62.5%=90点)。
#     ◎は従来どおり連対確率(2着以内に来る確率)1位。○▲△ は v136_010 の方式
#     (◎との馬連確率 × 妙味、λ=2)。穴 = 補正前人気4〜9位で 単勝確率がオッズ勝率の1.21倍以上・複勝確率25%以上。
#
#   【勝負レース】(本線枠・次点本線枠・統計裏付けは廃止。実弾の買い目は出さない)
#     ◎の単指数が90以上(単勝確率62.5%以上)のレースを『勝負レース』として表示する(2026/09/17〜29の506Rで約8%)。
#     紐(○▲△)の選び方: 勝負レース=単勝確率の高い順3頭 / それ以外のレース=期待値(単勝確率×予想オッズ)の高い順3頭。
#       (506Rの検証: 単指数90以上は単勝確率順が現行より有利、それ以外は期待値順が有利。REF_PARTNER_MODE=legacy で従来の妙味重視(λ=2)に戻る)
#     買い目は ◎→○▲△ の馬連・ワイド・馬単各3点(表示のみ。馬単は◎1着→紐2着)。勝負レース以外はワイドを表示しない。統計予想(統計◎○▲△・参考買い目)は従来どおり表示。
#     軸級・信頼度・レース分類は ◎ の単勝確率/複勝確率で判定。
#
#   【出力列の置き換え】(列の数・並びは v136_021 と同じ)
#     研究用HTML: 騎手補正→騎手指数 / (旧)単指数の列→前売(補正前単勝オッズ) / (旧)複指数の列→人気O(人気オッズ) /
#                 融合%→オッズ%(オッズ勝率)。単オッズ=予想オッズ、人気=予想オッズの順。脚質は出馬表HTMLから取得(取れない馬は『-』)。
#     Excel/CSV も同じ置き換え。keiba_tool008 で列名を見ている場合は読み替えが必要。
#     版情報 meta: bado-version=v137_001 / bado-logic に prob=odds+stat;shobu=idx90。
#
#   ※紐馬指数・騎手指数の重み・予想オッズの配合は、的中データでの検証をまだしていない初期値。
#     結果がたまったら検証すること。損失許容の範囲で。
#
#   (v136_021 以前の変更履歴は keiba_yosou_v136_021.py を参照)
# =====================================================

import argparse
import threading
from datetime import date, datetime
from pathlib import Path
from typing import NamedTuple

import pandas as pd
import numpy as np
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import json as _json
import os as _os
# ★v136_003: 旧・参考枠モジュール(gz_reference.py)は全廃した。
#   参考予想は generate_analysis 内の _build_reference_himo に一本化している。

# ════════════════════════════════════════════════════════════════
# ★ horselist(当日の *_horselist.csv)と読込中のファイル情報
#   (競馬場, 競走年月日, レース番号, 馬番) で各馬に結合する(read_racecard_csv)。
# ════════════════════════════════════════════════════════════════
_HORSELIST_TABLE_CACHE = {}   # folder(str) -> table(dict)
_HL_MULTIDATE_WARNED = False  # 複数日警告を1回だけ出す
# 予想実行時に「今読んでいるCSVのフォルダ」を入れる(batch/GUI/main で設定)
CURRENT_RACECARD_DIR = None
# 予想対象日(YYYYMMDD int)。レースカードのファイル名から推定。
# 複数日分の *_horselist.csv がある場合、この日付で厳密に絞り込む。
CURRENT_RACECARD_DATE = None

# ══════════════════════════════════════════════════════════════════
# ★v135_004: 均等買い(フラットベット)設定
#   Kelly機構(条件別の賭け金配分)は撤去した。理由:
#   ・prob_model_conditions_with_axis.json の invest は全12,858件が
#     n_bets×100 であり、学習/検証ROIは『1点100円の均等買い』で算出された
#     数字である。条件別のKelly重み付けは検証されていない後付けの操作。
#   ・旧実装は「条件ごとに1点分の金額」を全買い目に割り当てており、cap(6%)も
#     条件単位でしか効かなかったため、点数分だけ過剰にベットしていた
#     (実測: U2単独発火で意図2,480円に対し実投下14,400円=約5.8倍)。
#   ・同一軸から出る馬単複数点は排反事象であり、1点分のKellyをN点へ
#     そのまま掛けるのは理論的にも誤り。
#   → 1点100円固定 + 1レースあたりの投下上限(資金比)で管理する。
#
#   GZ_BANKROLL       総資金(円)              環境変数で上書き可
#   GZ_FLAT_UNIT      1点あたりの購入額(円)
#   GZ_RACE_CAP_RATIO 1レース投下上限の資金比(0.005 = 0.5%)
#   GZ_MAX_POINTS     1レース最大点数。0以下なら資金比から自動算出
# ══════════════════════════════════════════════════════════════════
import os as _os_gz


def _gz_env_num(key, default, cast=float):
    try:
        return cast(float(_os_gz.environ.get(key, str(default))))
    except Exception:
        return cast(default)


GZ_BANKROLL       = _gz_env_num('GZ_BANKROLL', 500000, int)
GZ_FLAT_UNIT      = _gz_env_num('GZ_FLAT_UNIT', 100, int)
GZ_RACE_CAP_RATIO = _gz_env_num('GZ_RACE_CAP_RATIO', 0.005, float)
GZ_MAX_POINTS     = _gz_env_num('GZ_MAX_POINTS', 0, int)

# ══════════════════════════════════════════════════════════════════
# ★v137_001: アンサンブル確率・予想オッズ・指数の設定(環境変数で上書き可)
#   ※いずれも的中データでの検証前の初期値。結果がたまったら見直すこと。
# ══════════════════════════════════════════════════════════════════
# オッズ勝率 ∝ (1/補正前オッズ)^ODDS_PROB_POWER。1.0=そのまま正規化 / >1 で人気薄をより割り引く。
ODDS_PROB_POWER = _gz_env_num('ODDS_PROB_POWER', 1.0, float)
# 単勝確率 ∝ オッズ勝率^ENS_W_ODDS × 統計勝率^ENS_W_STAT(対数線形の融合)
#   ★v137_001 最適化: 2026/09/17〜29の506R・5,178頭で勝ち馬の対数損失を最小化した値。
#     旧 0.70/0.30(和=1)は本命を低く・大穴を高く見積もる(30〜40%帯の予測34.7%に対し実際44.8%)。
#     統計比率6通り×合計4通りを総当たりし、統計比率0.3・合計1.15(=0.805/0.345)が的中率と回収率のバランス最良:
#     ◎単勝的中42.5%(最高)・紐3点 馬連 的中42.1%/回収97.0%・馬単 回収102.2%、前後半のぶれ最小。損失1.6729→1.6577。
#     (0.75/0.50 は損失最小1.6505だが紐の回収率が低く、◎の的中も下がった)
ENS_W_ODDS = _gz_env_num('ENS_W_ODDS', 0.805, float)
ENS_W_STAT = _gz_env_num('ENS_W_STAT', 0.345, float)
# 統計勝率が無いレースはオッズ勝率だけで予想する。同じ検証での最適な尖り(オッズ勝率^1.10)。
ENS_ODDS_ONLY_POWER = _gz_env_num('ENS_ODDS_ONLY_POWER', 1.10, float)
# 2着・3着の割引(Harville の拡張。Lo & Bacon-Shone の推定値に近い値)
PLACE_LAMBDA2 = _gz_env_num('PLACE_LAMBDA2', 0.81, float)
PLACE_LAMBDA3 = _gz_env_num('PLACE_LAMBDA3', 0.65, float)
# 払戻率(地方競馬の標準)。予想オッズ・馬連/ワイドの推定配当に使う。
TAKEOUT_WIN = _gz_env_num('TAKEOUT_WIN', 0.80, float)
TAKEOUT_PAIR = _gz_env_num('TAKEOUT_PAIR', 0.75, float)
# 予想オッズ = exp( W_RAW·log(補正前オッズ) + (1-W_RAW)·log(人気オッズ) )
POP_ODDS_W_RAW = _gz_env_num('POP_ODDS_W_RAW', 0.70, float)
# 人気オッズ(人気になりそうな項目の条件付きロジット係数)。2026/09/29 NAR 59R の市場オッズに当てはめた値。
#   jloo : 騎手の他の騎乗馬の人気(本馬を除く, log(オッズ勝率×頭数) の縮約平均)
#   jride: 騎手のその場の騎乗数(log, 場の平均との差)
#   app  : 減量騎手(horselist の負担重量に ☆▲△◇★)
#   lw   : 全成績の勝率(経験ベイズで縮約, logit)
#   lv   : 当地(当競馬場)の複勝率(縮約, logit)
#   lc   : この騎手とのコンビの複勝率(縮約, logit)
#   exp  : キャリア log(出走数+1)
#   bt   : 最高タイムのレース内偏差(速いほど+) / bt_na: 最高タイムなし
POP_COEF = dict(jloo=0.350, jride=0.297, app=-0.186, lw=0.145, lv=0.462, lc=0.242,
                exp=-0.414, bt=0.290, bt_na=0.323)
JOCKEY_LOO_SHRINK = 2.0          # 騎手の他の騎乗馬の人気を 0(平均)へ縮める強さ(騎乗数換算)
# 紐馬指数の配当の重み(2〜3着に来る確率 × 予想オッズ^HIMO_ODDS_POWER)
#   ★v137_001: 旧式(×(複勝確率÷市場の複勝確率)^0.5)は比がほぼ1で妙味が効かず、単勝確率順とほぼ同じだった。
#   506Rで ◎→紐馬指数上位3頭 の回収率 馬連78.8→88.0%・馬単77.7→88.8%(1位馬の複勝的中は50→44%)。
HIMO_ODDS_POWER = _gz_env_num('HIMO_ODDS_POWER', 0.5, float)
# 穴印: 補正前人気がこの範囲 & 単勝確率/オッズ勝率 >= ANA_RATIO_MIN & 複勝確率 >= ANA_PLACE_MIN
ANA_POP_RANGE = (4, 9)
ANA_RATIO_MIN = 1.21      # ★v137_001: 尖らせ後も旧(1.30)と同じ馬が選ばれるよう調整
ANA_PLACE_MIN = 25.0
KISHU_DEV_COL   = '騎手指数'      # ★v137_001: 予想表の騎手の列(新ロジックの騎手指数)
KISHU_DEV_LABEL = '騎手指数'

# ★v135_015: 紐カバー枠(tier='cover')は本線枠とは別勘定にする。
#   本線枠の点数上限を穴条件が食い潰すと、的中率の高い条件が
#   「点数上限で見送り」になり本末転倒なので、財布を分ける。
#   既定は本線枠の 2/5 (0.002 = 資金の0.2%)。0にすれば穴枠は完全停止。
GZ_ANA_CAP_RATIO  = _gz_env_num('GZ_ANA_CAP_RATIO', 0.002, float)
GZ_ANA_MAX_POINTS = _gz_env_num('GZ_ANA_MAX_POINTS', 0, int)


def gz_ana_max_points():
    """紐カバー枠に購入できる最大点数(本線枠とは独立)。"""
    if GZ_ANA_MAX_POINTS > 0:
        return GZ_ANA_MAX_POINTS
    if GZ_FLAT_UNIT <= 0 or GZ_ANA_CAP_RATIO <= 0:
        return 0
    return int(GZ_BANKROLL * GZ_ANA_CAP_RATIO // GZ_FLAT_UNIT)


def gz_race_max_points():
    """1レースに購入できる最大点数。GZ_MAX_POINTS>0 ならそれを優先。

    ★v135_010: 以前は max(1, ...) で最低1点を保証していたため、資金が
      2万円を下回ると 1点=100円 が資金比0.5%を超えていた(1万円なら1.0%)。
      上限を名乗る以上は超えないのが正しいので、上限額が1点分に満たない
      資金では 0点(=そのレースは見送り)を返す。
    """
    if GZ_MAX_POINTS > 0:
        return GZ_MAX_POINTS
    if GZ_FLAT_UNIT <= 0:
        return 0
    return int(GZ_BANKROLL * GZ_RACE_CAP_RATIO // GZ_FLAT_UNIT)

def _guess_date_from_name(path):
    '''ファイル名から YYYYMMDD を推定。
    例) race_data001_updatedtihou0712.csv → 当年+0712 を優先。
        20260712_... のような8桁があればそれを最優先。'''
    import os as _os, re as _re
    from datetime import date as _date
    base = _os.path.basename(str(path))
    # (a) 8桁 YYYYMMDD
    m8 = _re.findall(r'(20\d{6})', base)
    if m8:
        try:
            return int(m8[-1])
        except Exception:
            pass
    # (b) 4桁 MMDD(月日のみ) → 実行日に最も近い年の同月日を採用。
    #   前年/当年/翌年の3候補のうち、実行日との差(絶対値)が最小のものを選ぶ。
    #   これにより 12/30実行→0103(翌年頭)や 1/3実行→1230(前年末)の
    #   年跨ぎ予想も、年に依存せず正しい側へ寄る。
    #   ※最終的には _attach_hl_to_group で「推定日が horselist に実在する
    #     場合のみ採用」するため、多少ずれても実データで救済される。
    m4 = _re.findall(r'(?<!\d)(\d{4})(?!\d)', base)
    _t = _date.today()
    _t_ord = _t.toordinal()
    for tok in reversed(m4):
        mm, dd = int(tok[:2]), int(tok[2:])
        if 1 <= mm <= 12 and 1 <= dd <= 31:
            best = None
            for y in (_t.year - 1, _t.year, _t.year + 1):
                try:
                    diff = abs(_date(y, mm, dd).toordinal() - _t_ord)
                except ValueError:
                    continue   # 2/29 が無い年など
                if best is None or diff < best[0]:
                    best = (diff, y * 10000 + mm * 100 + dd)
            if best is not None:
                return best[1]
    return None

# 全角数字→半角(レース番号 '１'→'1' 等。CSVによって全角/半角が混在するため)
_Z2H_DIGITS = str.maketrans('０１２３４５６７８９', '0123456789')

# レースのグルーピングキー(出力順序を決める)
_RACE_GROUP_KEYS = ['場所', '距離', 'レース', '出走時刻']



def _sort_by_race_no(df):
    """レース番号で場所別に昇順ソートし、'レース数'列を付与して返す。
    ・文字列順だと '1R'→'10R'→'11R'→'2R' になるため数値抽出してから並べる。
      groupby(sort=False) がこの並び順を保持するので Excel/HTML の出力順も正しくなる。
    ・'１０R' のような全角数字が来るので先に半角化する
      (全角のままだと to_numeric が NaN→0 になりソートが効かない)。
    ・数字を含まないレース名は 0 扱い(astype(int) で落ちないように coerce)。
    """
    df['レース'] = df['レース'].astype(str)
    df['レース数'] = pd.to_numeric(
        df['レース'].str.translate(_Z2H_DIGITS).str.extract(r'([0-9]+)', expand=False),
        errors='coerce').fillna(0).astype(int)
    return df.sort_values(by=['場所', 'レース数'], kind='mergesort')

def _to_int_z2h(v):
    """全角/半角どちらの数字文字列でも int に変換。不能なら None。"""
    try:
        import pandas as _pd
        s = str(v).strip().translate(_Z2H_DIGITS)
        n = _pd.to_numeric(s, errors='coerce')
        if n != n:   # NaN
            return None
        return int(n)
    except Exception:
        return None

def _race_no_of(v):
    """'1R' / '１０Ｒ' / 3 → レース番号 int。_to_int_z2h は 'R' 付きだと None になるため別に用意。"""
    m = re.search(r'\d+', unicodedata.normalize('NFKC', str(v)))
    return int(m.group()) if m else v


def _norm_venue_hl(v):
    s = str(v).strip()
    if s.endswith('ば') and len(s) > 1:
        s = s[:-1]
    return s

_HL_MISSING_WARNED = False   # horselist不在警告を1回だけ出す


def _read_csv_auto(path, dtype=str):
    """cp932 / utf-8-sig / utf-8 を順に試して読む(列名の BOM は除く)。"""
    last = None
    for enc in ('utf-8-sig', 'cp932', 'utf-8', 'shift_jis'):
        try:
            df = pd.read_csv(path, encoding=enc, dtype=dtype)
            df.columns = [str(c).lstrip('﻿').strip() for c in df.columns]
            return df
        except Exception as e:
            last = e
    raise last


def _get_horselist_table(folder):
    """★v137_001: フォルダ内の *_horselist.csv を直接読み、(競馬場, 日付, R, 馬番) -> 行dict を返す。
    (v136_021 までの horselist_features.py は不要。値は文字列のまま。)"""
    global _HL_MISSING_WARNED
    if not folder:
        return {}
    folder = str(folder)
    try:
        import glob as _glob_hl
        _files = sorted(_glob_hl.glob(_os_gz.path.join(folder, '*_horselist.csv')))
        _sig = tuple((_p, _os_gz.path.getmtime(_p), _os_gz.path.getsize(_p)) for _p in _files)
    except Exception:
        _files, _sig = [], ()
    _ck = (folder, _sig)
    if _ck in _HORSELIST_TABLE_CACHE:
        return _HORSELIST_TABLE_CACHE[_ck]
    for _k in [k for k in _HORSELIST_TABLE_CACHE if isinstance(k, tuple) and k[0] == folder]:
        del _HORSELIST_TABLE_CACHE[_k]
    table = {}
    for _p in _files:
        try:
            _df = _read_csv_auto(_p)
        except Exception as _e:
            print(f'  [horselist警告] {_p} を読めません: {_e}')
            continue
        _need = ('競馬場', '競走年月日', 'レース番号', '馬番')
        if not all(c in _df.columns for c in _need):
            print(f'  [horselist警告] {_p} に必要な列がありません: {[c for c in _need if c not in _df.columns]}')
            continue
        for _r in _df.to_dict('records'):
            _d = _to_int_z2h(_r.get('競走年月日'))
            _rn = _to_int_z2h(_r.get('レース番号'))
            _u = _to_int_z2h(_r.get('馬番'))
            if _d is None or _rn is None or _u is None:
                continue
            table[(_norm_venue_hl(_r.get('競馬場')), _d, _rn, _u)] = _r
    if not table and not _HL_MISSING_WARNED:
        print(f'  [horselist警告] {folder} に読める *_horselist.csv が見つかりません。\n'
              f'    オッズCSVと同じフォルダに当日の horselist を置いてください。\n'
              f'    このままでも動作しますが、騎手変更・人気オッズの材料(成績/最高タイム)が欠けます。')
        _HL_MISSING_WARNED = True
    _HORSELIST_TABLE_CACHE[_ck] = table
    return table


# ─────────────────────────────────────────────────────────────
# NamedTuple 定義
# ─────────────────────────────────────────────────────────────
class SingleRec(NamedTuple):
    ban: int
    hit_rate: float
    odds: float
    ev: float


class PairRec(NamedTuple):
    ban1: int
    ban2: int
    hit_rate: float
    odds: float
    ev: float
    compat: float
    sort_score: float


# =====================================================
# Config
# =====================================================
class Config:
    import os as _os
    DEFAULT_DATA_PATH    = Path(_os.environ.get('RACE_DATA_PATH', 'race_data.csv'))
    DEFAULT_OUTPUT_DIR   = Path(_os.environ.get('RACE_OUTPUT_DIR', '.'))
    # ★v137_001: 旧版の投資サマリー・騎手オッズ補正・ワイド係数は不使用(払戻率は TAKEOUT_* を参照)


def _parse_args():
    parser = argparse.ArgumentParser(description='地方競馬予想 v137_001 オッズ×統計 アンサンブル版')
    parser.add_argument('--data',   type=Path, default=Config.DEFAULT_DATA_PATH,
                        help='単勝オッズCSV(nar_shutuba_all_YYYYMMDD.csv)')
    parser.add_argument('--output', type=Path, default=Config.DEFAULT_OUTPUT_DIR)
    parser.add_argument('--deba', type=Path, default=None,
                        help='★v136_019 出馬表HTMLフォルダ(起動時に読み込む)')
    args, _ = parser.parse_known_args()
    return args

_args = None



# =====================================================
# ★ v130: 的中信頼度インジケータ（オススメ度を「的中しやすさ」指標へ）
#   調査結論: 購入区分のROI単調化は汎化しない(4/3/2区分すべて並べ替え検定で非有意)一方、
#   的中率は自信度と単調に対応し汎化する。よってオススメ度は
#   「レースの推奨買い目が的中しやすいか」を表す信頼度(0-100)として算出する。
#   ・軸の複勝確率/勝率を核に、支配度(z)とNCS級で微調整(埋め込み既定)。
#   ・同フォルダ or CONFIDENCE_FORMULA_PATH に confidence_formula.json があれば
#     バックテスト(target=hit, ウォークフォワード検証済)の重み・較正・区分境界を自動採用。
# =====================================================
_CONF_FORMULA = None   # 外部フィット式(あれば)

_CONF_DEFAULT = dict(
    w_place=0.62, w_win=0.38,           # 0-100の確率ブレンド(複勝主導)
    z_gain=4.0, z_clip=[-10.0, 12.0],   # 軸の支配度(z-score)による微調整
    grade_adj={'S': 6.0, 'A': 3.0, 'B': 0.0, 'C': -3.0, 'D': -6.0},
)
_CONF_TIERS = [
    (78, '非常に高い', 'FFFFFF', '375623', 'C6EFCE'),
    (62, '高い',       'FFFFFF', '1F4E79', 'BDD7EE'),
    (46, '中程度',     '000000', 'FFEB9C', 'FFF2CC'),
    (30, '低い',       'FFFFFF', 'ED7D31', 'F8CBAD'),
    (-1, '非常に低い', 'FFFFFF', 'C00000', 'FFC7CE'),
]


def _load_confidence_formula():
    """外部 confidence_formula.json があれば読み込む(無ければ埋め込み既定を使用)。"""
    import os as _os4
    global _CONF_FORMULA
    cands = []
    envp = _os4.environ.get('CONFIDENCE_FORMULA_PATH')
    if envp:
        cands.append(envp)
    try:
        cands.append(_os4.path.join(_os4.path.dirname(_os4.path.abspath(__file__)), 'confidence_formula.json'))
    except Exception:
        pass
    cands.append('confidence_formula.json')
    for p in cands:
        try:
            if p and _os4.path.exists(p):
                with open(p, encoding='utf-8') as f:
                    _CONF_FORMULA = _json.load(f)
                _CONF_FORMULA['_src'] = p
                return p
        except Exception:
            continue
    return None
_load_confidence_formula()


def _conf_tier(score, cutoffs=None):
    if cutoffs:
        for row in cutoffs:
            label, fh, bg, bar, mn = row
            if score >= mn:
                return label, fh, bg, bar
        row = cutoffs[-1]
        return row[0], row[1], row[2], row[3]
    for mn, label, fh, bg, bar in _CONF_TIERS:
        if score >= mn:
            return label, fh, bg, bar
    return _CONF_TIERS[-1][1:]


def _conf_apply_formula(analysis, F, place, win, z, grade):
    """外部フィット式(v2: 確率ブレンド)で信頼度(0-100)と推定的中率を計算。
    v2は place/win(0-100の較正済み確率)への凸ブレンドなので、学習時と予想時で
    スケールが必ず一致し、崩壊しない。旧v1(standardize+percentile)は無視。"""
    blend = F.get('blend')
    if not blend:
        return None, None            # v2でなければ既定にフォールバック
    comp = float(analysis.get('axis_analysis', 0) or 0)
    raw = (float(blend.get('複勝確率', 0.0)) * place +
           float(blend.get('単勝確率', 0.0)) * win +
           float(blend.get('複合スコア', 0.0)) * min(comp, 100.0))
    zc = F.get('z_clip', _CONF_DEFAULT['z_clip'])
    z_adj = float(np.clip(z * float(F.get('z_gain', _CONF_DEFAULT['z_gain'])), zc[0], zc[1]))
    gadj = F.get('grade_adj', _CONF_DEFAULT['grade_adj'])
    conf = raw + z_adj + float(gadj.get(grade, 0.0))
    conf = max(0.0, min(100.0, conf))
    est_hit = None
    hbc = F.get('hit_by_conf')        # 長さ101: conf 0..100 → 実的中率%
    if hbc:
        est_hit = float(hbc[int(round(conf))])
    return conf, est_hit


# ════════════════════════════════════════════════════════════════
# ★v136_008: 新判定(軸級・信頼度・レース分類) ── ◎の推定確率ベース
#   ★v136_009: 係数を 2026/07/16〜09/19 の NAR 2,324R で再当てはめ
#     (予想表の単勝確率・補正後単勝オッズ＋配当CSV。horselist結合失敗で
#      モデル確率が崩れた 7/22・8/6・8/9・8/25 の169Rは除外)。
#   検証: 7〜8月で当てはめ→9月(779R, 当てはめに未使用)でも各区分が単調・較正。
#     3着内率の式に ◎勝率・頭数を足す案は前向き検証/ブロックCVで改善せず不採用。
#   (v136_008 の係数: A=0.391 B=0.779 SLOPE=0.6649 INTERCEPT=-0.1595 / 1,714R)
# ════════════════════════════════════════════════════════════════
USE_NEW_JUDGE = bool(_gz_env_num('USE_NEW_JUDGE', 1, int))   # 0 で旧判定に戻す

# ブレンド勝率: p ∝ (単勝確率)^A × (オッズ由来確率)^B をレース内で正規化
#   ★v137_001: 単勝確率がすでにオッズとのアンサンブルなので A=1, B=0(=単勝確率そのもの)。
#     ref_partner_order の妙味(オッズ由来確率との比)には補正前オッズを渡す。
JUDGE_BLEND_A = 1.0
JUDGE_BLEND_B = 0.0
JUDGE_ODDS_MISSING = 199.0      # オッズ欠損馬は最低人気相当として扱う
# 推定3着内率: logit(p3) = SLOPE × logit(Harville3着内確率) + INTERCEPT
JUDGE_TOP3_SLOPE = 0.6374
JUDGE_TOP3_INTERCEPT = -0.1356
# 軸級: ◎の推定勝率(%)の下限。これ未満は D級
#   実績(7/16〜9/19 2,324R): S 勝率62.0%/複勝圏88.1%, A 42.6/77.6, B 31.2/67.0, C 27.8/62.4, D 21.1/49.9
# ★v137_001: 単勝確率の尖らせ(0.805/0.345)に合わせ、旧確率と同じ構成比になるよう分位で再調整。
#   旧 S50/A40/B31/C24 → S56/A44.5/B34.3/C26.1(S12.3%・A16.4%・B27.0%・C24.3%の構成比を維持)。
JUDGE_GRADE_CUTS = [('S', 56.0), ('A', 44.5), ('B', 34.3), ('C', 26.1)]
# 信頼度: 数値 = ◎の推定3着内率(%)。区分境界(色は旧区分と同じ)
#   実績(7/16〜9/19 2,324R・旧確率): 非常に高い 88.6%, 高い 78.5, 中程度 68.2, 低い 59.4, 非常に低い 48.5
#   ★v137_001: 尖らせ後の複勝確率に合わせて 83/74/64/55 → 88/79/68/59 (旧と同じ構成比)
JUDGE_CONF_TIERS = [
    (88, '非常に高い', 'FFFFFF', '375623', 'C6EFCE'),
    (79, '高い',       'FFFFFF', '1F4E79', 'BDD7EE'),
    (68, '中程度',     '000000', 'FFEB9C', 'FFF2CC'),
    (59, '低い',       'FFFFFF', 'ED7D31', 'F8CBAD'),
    (-1, '非常に低い', 'FFFFFF', 'C00000', 'FFC7CE'),
]
# レース分類(上から順に判定)
#   ★v137_001: 尖らせ後の確率に合わせて再調整(旧 60/88/76/55 と同じ構成比: 投資6.3%・有力+約27%)
JUDGE_CAT_INVEST_WIN = 66.5     # 投資: 推定勝率66.5%以上 かつ
JUDGE_CAT_INVEST_TOP3 = 92.0    #       推定3着内率92%以上
JUDGE_CAT_STRONG_TOP3 = 81.0    # 有力: 推定3着内率81%以上
# ★v136_009: 2頭軸の判定は停止。9月(未使用データ)で ◎-対抗 馬連的中 1/20R(予測35%)、
#   7/16〜9/19 通算でも 26%(85R)と有力レース(約30%)を下回り、区分として機能しなかった。
#   該当していたレースは 標準レース(推定3着内率55%未満なら見送り=表示は標準) になる。
JUDGE_CAT_TWO_ENABLE = False    # True で再開(下の2つの閾値を使用)
JUDGE_CAT_TWO_Q = 30.0          # 2頭軸: ◎-対抗の馬連確率30%以上 かつ
JUDGE_CAT_TWO_RIVAL = 25.0      #        対抗のブレンド勝率25%以上
JUDGE_CAT_SKIP_TOP3 = 59.0      # 見送り: 推定3着内率59%未満(表示は _RACE_CAT_MASK で標準レース)


def _judge_harville_top3(p, i):
    """Harville モデルで馬 i の3着内確率を返す(p はレース内合計1の勝率配列)。"""
    n = len(p)
    pi = p[i]
    tot = pi
    for j in range(n):
        if j == i:
            continue
        pj = p[j]
        d1 = 1.0 - pj
        if d1 <= 1e-12:
            continue
        tot += pj * pi / d1
        for k in range(n):
            if k == i or k == j:
                continue
            d2 = 1.0 - pj - p[k]
            if d2 <= 1e-12:
                continue
            tot += pj * p[k] / d1 * pi / d2
    return float(min(max(tot, 0.0), 1.0))


def _judge_blend(win_pct, adj_odds):
    """(pm, pk, p) を返す。pm=モデル単勝確率, pk=オッズ由来確率, p=ブレンド勝率(いずれも合計1)。
    win_pct : 予想表の単勝確率(%, レース内合計≒100)
    adj_odds: 予想表の単勝オッズ(騎手補正後 adjusted_odds)。0以下/欠損は欠損扱い。"""
    pm = np.asarray(pd.to_numeric(pd.Series(win_pct), errors='coerce').fillna(0.0), dtype=float)
    pm = np.clip(pm, 0.1, None)
    pm = pm / pm.sum()
    od = np.asarray(pd.to_numeric(pd.Series(adj_odds), errors='coerce').fillna(0.0), dtype=float)
    od = np.where(od > 0, np.clip(od, 1.0, None), JUDGE_ODDS_MISSING)
    if np.all(od >= JUDGE_ODDS_MISSING):
        pk = pm.copy()                       # オッズ全欠損 → モデルのみ
    else:
        pk = (1.0 / od) / np.sum(1.0 / od)
    s = np.exp(JUDGE_BLEND_A * np.log(pm) + JUDGE_BLEND_B * np.log(pk))
    p = s / s.sum()
    return pm, pk, p


# ════════════════════════════════════════════════════════════════
# ★v136_010: 予想印○▲△ と 参考買い目(◎から3点)の相手選定
#   相手 = ◎以外の馬を『◎との馬連確率(ブレンド勝率のHarville) × 妙味』の順に3頭。
#     スコア = log Q_blend + λ・log(Q_blend / Q_odds)
#       Q_blend: ブレンド勝率から見た ◎-相手 の馬連確率(当たりやすさ)
#       Q_odds : 単勝オッズ由来確率から見た同じ組の馬連確率(≒配当の安さ)
#       λ=1 は「当たりやすさ × 配当の割安さ」を等しく重視する。
#   検証(2026/07/16〜09/19 2,493R, ◎・3点固定, 1点100円):
#     旧(○▲△☆の先頭3頭)  馬連 的中35.3% 回収74.6% / ワイド 的中57.0% 回収76.5%
#     新(λ=1)              馬連 的中41.6% 回収78.2% / ワイド 的中61.7% 回収76.9%
#     的中率の改善は 馬連+6.3pt[95%CI +4.9,+7.7] ワイド+4.7pt[+3.4,+5.9]。
#     回収率の差は +3.6pt / +0.3pt で誤差の範囲(悪化はしていない)。
#   λ=0(ブレンド勝率順)は的中率は同等だが回収率が下がり、λ=3 やモデル単勝率順は
#   回収率がやや高い代わりに的中率の改善が半分程度だったため、λ=1 を採用。
#   ★v136_015: 回収率重視に切り替え λ=4(ワイド回収 76.9→81.2%, 馬連 78.2→79.8%)。
#   REF_NEW_LOGIC=0 で旧ロジック(不具合修正済み)に戻せる。
# ════════════════════════════════════════════════════════════════
REF_NEW_LOGIC = bool(_gz_env_num('REF_NEW_LOGIC', 1, int))

# ★v136_014: 予想表に埋め込む版情報(keiba_tool008 が読む)。版を上げたらここも更新する。
BADO_VERSION = 'v137_001'


def bado_version_meta():
    """予想表HTMLの <head> に入れる版情報の meta タグ(2本)を返す。"""
    _logic = 'prob=odds+stat;judge=%s;ref=%s;shobu=idx90;partner=%s;stat=%s;obs=off;dq=1' % (
        'new' if USE_NEW_JUDGE else 'old', 'new' if REF_NEW_LOGIC else 'old', REF_PARTNER_MODE,
        'stat1' if (STAT_DEBA_RACES and stat_model()) else 'off')
    return ('<meta name="bado-version" content="%s">'
            '<meta name="bado-logic" content="%s">' % (BADO_VERSION, _logic))


# ══════════════════════════════════════════════════════════════════
# ★v136_019: 統計予想(出馬表HTML × 既存データの融合)
#   bado_stat_model.py / bado_stat_model.json を同じフォルダに置く。無ければ統計列は空のまま。
# ══════════════════════════════════════════════════════════════════
import re
import unicodedata
try:
    import bado_stat_model as _bsm
except Exception as _e_bsm:
    _bsm = None
    print(f'[統計予想] bado_stat_model.py を読み込めません({_e_bsm})。統計予想は無効です。')
STAT_ENABLE = True
STAT_MODEL = None            # None=未読込 / False=読込失敗 / dict
STAT_DEBA_RACES = {}         # {(日付, 競馬場, R): DebaRace}
STAT_DEBA_FOLDER = ''
STAT_TARGET_DATE = None      # 予想対象CSVの日付(YYYYMMDD)。不明なら None
STAT_NAME_MATCH_MIN = 0.7    # 馬番→馬名の一致率がこれ未満のレースは結合しない(日付・場の取り違え防止)
STAT_MARKS = ('◎', '○', '▲', '△')
# ★v136_020: 統計裏付け(本線/次点の軸馬の統計勝率がこれ以上)
STAT_URA_MIN = 50.0
STAT_URA_MARK = '裏'
STAT_URA_LABEL = '統計裏付け'
# 本線/次点の条件コード → 軸の種類(★v137_001)
STAT_AXIS_OF = {}
STAT_AXIS_CODES = ()
# 軸の種類 → generate_analysis の軸キー
STAT_AXIS_KEY = {}
# ★v136_020 の観察条件 ST1〜ST6 は旧軸((旧)外部AIの単指数)に依存するため停止(空)。
STAT_OBS_DEFS = []


def _stat_ura_of(analysis, code):
    """条件コード(EU1/EU2/EW1/XW1/XU1)の軸に統計裏付けがあれば {ban, ps, ok} を返す。"""
    try:
        _u = (analysis.get('stat_ura') or {}).get(STAT_AXIS_OF.get(str(code), ''))
        return _u if (_u and _u.get('ok')) else None
    except Exception:
        return None


# ══════════════════════════════════════════════════════════════════
# ★v136_021: データ確認(horselist結合率・騎手変更・出走取消・分析エラー)
# ══════════════════════════════════════════════════════════════════
HL_WARN_RATE = 0.90          # これ未満なら全体警告
HL_RACE_WARN_RATE = 0.50     # レース単位でこれ未満なら注意書き
JK_CHANGE_MARK = '※'
RUN_ERRORS = []              # _run_output 中に分析できなかったレース
_JK_VARIANTS = str.maketrans({'濱': '浜', '髙': '高', '﨑': '崎', '邊': '辺', '邉': '辺', '齋': '斉', '齊': '斉', '斎': '斉',
                              '澤': '沢', '廣': '広', '德': '徳', '眞': '真', '惠': '恵', '國': '国', '櫻': '桜',
                              '龍': '竜', '嶋': '島', '嶌': '島', '冨': '富', '瀨': '瀬'})


def _jk_nz(s):
    s = unicodedata.normalize('NFKC', str(s or ''))
    s = re.sub(r'[\s▲△☆◇★*※]', '', s)
    s = re.sub(r'^[A-Za-z]+\.', '', s)          # 外国人騎手の頭文字(Ｆ．ゴン → ゴン)
    return s.translate(_JK_VARIANTS)


def _jockey_same(a, b):
    """オッズCSVとhorselistで騎手の略し方が違う(例 小笠原/小笠羚・多田羅/多田誠)ので、先頭2文字が同じなら同一騎手。"""
    na, nb = _jk_nz(a), _jk_nz(b)
    if not na or not nb or na == nb:
        return True
    return len(na) >= 2 and len(nb) >= 2 and na[:2] == nb[:2]



def stat_scratched_for_group(group):
    """出馬表HTMLで出走取消・競走除外の馬 {馬番: 馬名}。HTML未読込なら空。"""
    if not (STAT_DEBA_RACES and _bsm is not None):
        return {}
    try:
        _venue = re.sub(r'\s', '', unicodedata.normalize('NFKC', str(group['場所'].iloc[0])))
        _rno = int(unicodedata.normalize('NFKC', str(group['レース'].iloc[0])).replace('R', '').strip())
        _names = {int(_to_int_z2h(b)): str(n) for b, n in zip(group['番'], group['馬名'])
                  if _to_int_z2h(b) is not None}
        _race = _stat_find_race(_venue, _rno, _names)
    except Exception:
        return {}
    if _race is None:
        return {}
    return {int(u): str(nm) for u, nm in (getattr(_race, 'scratched', None) or []) if int(u) in _names}


def _dq_text_lines(analysis):
    """レース単位のデータ注意(Excel/CSV/HTML共通の文言)。"""
    out = []
    for b, nm in sorted((analysis.get('dq_scratched') or {}).items()):
        out.append('出走取消/除外 %d番%s(出馬表HTML)→予想から除外' % (b, nm))
    for c in (analysis.get('dq_jk') or []):
        out.append('騎手変更 %s番%s %s→%s(horselist)・新騎手の騎手指数%.0f' % (c['ban'], c['name'], c['old'], c['new'],
                                                                   c.get('idx_new') or 50.0))
    _n, _m = analysis.get('hl_n') or 0, analysis.get('hl_m') or 0
    if _n and _m / _n < HL_RACE_WARN_RATE:
        out.append('horselist未結合 %d/%d頭 → 人気オッズ・騎手変更の材料が欠ける(当日のhorselistを確認)' % (_m, _n))
    if analysis.get('dq_nostat') and STAT_DEBA_RACES:
        out.append('統計勝率なし(出馬表HTMLと結合できない) → 単勝確率はオッズ勝率のみ')
    return out


def _data_note_html(analysis):
    _l = _dq_text_lines(analysis)
    if not _l:
        return ''
    return ('<div class="dnote">⚠ ' + ' ／ '.join(_h(x) for x in _l) + '</div>')


def data_quality_summary(analyses, errors=None, n_failed=0):
    analyses = [a for a in analyses if a]
    tot = sum(a.get('hl_n') or 0 for a in analyses)
    m = sum(a.get('hl_m') or 0 for a in analyses)
    rate = (m / tot) if tot else None
    low = [a.get('race_label', '') for a in analyses
           if (a.get('hl_n') or 0) and (a.get('hl_m') or 0) / a['hl_n'] < HL_RACE_WARN_RATE]
    jk = [(a.get('race_label', ''), c) for a in analyses for c in (a.get('dq_jk') or [])]
    scr = [(a.get('race_label', ''), b, nm) for a in analyses for b, nm in (a.get('dq_scratched') or {}).items()]
    errs = list(errors or [])
    nostat = [a.get('race_label', '') for a in analyses if a.get('dq_nostat')]
    return dict(tot=tot, m=m, rate=rate, low=low, jk=jk, scr=scr, errors=errs, n_failed=n_failed,
                nostat=nostat, nrace=len(analyses),
                severe=bool((rate is not None and rate < HL_WARN_RATE) or errs or n_failed))


def data_quality_lines(q):
    out = []
    if q['rate'] is not None and q['rate'] < HL_WARN_RATE:
        out.append('【重要】horselist結合率 %.0f%%(%d/%d頭)。当日の *_horselist.csv がオッズCSVと同じフォルダに無いか、'
                   '日付が合っていません。人気オッズ・騎手指数が不正確です。' % (q['rate'] * 100, q['m'], q['tot']))
        if q['low']:
            out.append('  未結合のレース: ' + ' '.join(q['low'][:12]) + (' ほか' if len(q['low']) > 12 else ''))
    elif q['low']:
        out.append('horselist未結合のレース: ' + ' '.join(q['low'][:12]))
    if q['errors'] or q['n_failed']:
        out.append('【重要】分析に失敗したレース %d件(予想表に出ていません):' % max(len(q['errors']), q['n_failed']))
        out.extend('  ' + e for e in q['errors'][:8])
    if q.get('nostat'):
        if not STAT_DEBA_RACES:
            out.append('【注意】出馬表HTMLが未読込のため統計勝率なし → 全レース オッズ勝率のみで予想'
                       '(『出馬表HTMLフォルダを読込』で読み込むか、オッズCSVと同じフォルダに R01_*.html を置く)。')
        elif not stat_model():
            out.append('【注意】bado_stat_model.py / .json が無いため統計勝率なし → オッズ勝率のみで予想。')
        else:
            out.append('統計勝率なし(出馬表HTMLと結合できない) %d/%dR: ' % (len(q['nostat']), q.get('nrace', 0))
                       + ' '.join(q['nostat'][:12]) + (' ほか' if len(q['nostat']) > 12 else ''))
    if q['jk']:
        out.append('騎手変更 %d件(horselistの騎手で計算):' % len(q['jk']))
        out.extend('  %s %s番%s %s→%s' % (r, c['ban'], c['name'], c['old'], c['new']) for r, c in q['jk'][:10])
    if q['scr']:
        out.append('出走取消/除外 %d頭(出馬表HTML・予想から除外): ' % len(q['scr'])
                   + ' '.join('%s %d番%s' % x for x in q['scr'][:10]))
    return out


def data_quality_banner_html(q):
    _l = data_quality_lines(q)
    if not _l:
        return ''
    _bg, _bd = ('#fde2e1', '#c0392b') if q['severe'] else ('#fff4d6', '#d4a017')
    return ('<div style="margin:8px 12px;padding:8px 12px;border:2px solid %s;background:%s;border-radius:6px;'
            'font-size:13px;line-height:1.6;color:#3a2a00"><b>データ確認</b><br>%s</div>'
            % (_bd, _bg, '<br>'.join(_h(x) for x in _l)))


def _stat_ura_badge(analysis, code):
    _u = _stat_ura_of(analysis, code)
    return (' <span class="ura" title="軸%d番の統計勝率%.1f%%">%s</span>' % (_u['ban'], _u['ps'], STAT_URA_LABEL)
            if _u else '')


def stat_model():
    global STAT_MODEL
    if STAT_MODEL is None:
        if _bsm is None:
            STAT_MODEL = False
        else:
            try:
                STAT_MODEL = _bsm.load_model()
            except Exception as _e:
                print(f'[統計予想] bado_stat_model.json を読み込めません({_e})。')
                STAT_MODEL = False
    return STAT_MODEL or None


def stat_load_deba_folder(folder):
    global STAT_DEBA_RACES, STAT_DEBA_FOLDER
    if _bsm is None:
        raise RuntimeError('bado_stat_model.py が予想コードと同じフォルダにありません。')
    STAT_DEBA_RACES = _bsm.load_deba_folder(folder)
    STAT_DEBA_FOLDER = str(folder)
    print(f'[統計予想] 出馬表HTML {len(STAT_DEBA_RACES)} レースを読み込みました: {stat_summary()}')
    return STAT_DEBA_RACES


def stat_set_target_date(date_str):
    global STAT_TARGET_DATE
    STAT_TARGET_DATE = date_str if (date_str and re.fullmatch(r'20\d{6}', str(date_str))) else None


def stat_set_target_date_from_name(name):
    # ★v136_021: horselist の対象日と同じ推定(年またぎ対応・月日の妥当性チェック)に統一
    try:
        _d = _guess_date_from_name(str(name))
    except Exception:
        _d = None
    stat_set_target_date(str(_d) if _d else None)


def stat_summary():
    if not STAT_DEBA_RACES:
        return '出馬表HTML 未読込'
    _v = {}
    for (_d, _ven, _r) in STAT_DEBA_RACES:
        _v.setdefault((_d, _ven), 0)
        _v[(_d, _ven)] += 1
    _ds = sorted({_d for _d, _ in _v})
    _vs = '・'.join(sorted({_ven for _, _ven in _v}))
    _dt = '/'.join(f'{int(_d[4:6])}/{int(_d[6:])}' for _d in _ds)
    return f'{len(STAT_DEBA_RACES)}R（{_vs}  {_dt}）'


def _stat_find_race(venue, rno, names_by_ban):
    _c = [(k, r) for k, r in STAT_DEBA_RACES.items() if k[1] == venue and k[2] == rno]
    if STAT_TARGET_DATE:
        _c = [(k, r) for k, r in _c if k[0] == STAT_TARGET_DATE]
    if not _c:
        return None
    _c.sort(key=lambda kr: kr[0][0], reverse=True)
    for _k, _r in _c:
        _hn = {h['uma']: _bsm._z(h['name']) for h in _r.horses}
        _both = [b for b in names_by_ban if b in _hn]
        if not _both:
            continue
        _ok = sum(1 for b in _both if _hn[b] == _bsm._z(names_by_ban[b])) / len(_both)
        if _ok >= STAT_NAME_MATCH_MIN:
            return _r
    return None



def _stat_td(row):
    # 研究用HTMLの馬表の最後の3列(統計印・統計%・オッズ%)  ★v137_001: 融合% → オッズ勝率
    _mk = str(row.get('統計印', '') or '')
    def _f(v):
        try:
            v = float(v)
            return f'{v:.1f}' if v == v else ''
        except Exception:
            return ''
    _cls = ' class="sthi"' if _mk.startswith('◎') else ''
    _mcls = ' class="sthi uracell"' if STAT_URA_MARK in _mk else _cls
    return f'<td{_mcls}>{_mk}</td><td>{_f(row.get("統計勝率"))}</td><td>{_f(row.get("オッズ勝率"))}</td>'


def _stat_text_lines(analysis):
    """★v136_020: Excel/CSV用。統計裏付けと観察条件(成立分)の説明行。"""
    _out = []
    _ura = analysis.get('stat_ura') or {}
    if _ura:
        _seen = {}
        for _k in STAT_AXIS_CODES:
            if _k in _ura:
                _seen.setdefault(_ura[_k]['ban'], []).append(_k)
        _out.append('軸の統計勝率(50%以上=統計裏付け): ' + ' / '.join(
            '%d番 %s 統計%.1f%%%s' % (_b, '・'.join(_ks), _ura[_ks[0]]['ps'],
                                     ' ' + STAT_URA_LABEL if _ura[_ks[0]]['ok'] else ' 裏付けなし')
            for _b, _ks in _seen.items()))
    _jp = {'umaren': '馬連', 'wide': 'ワイド'}
    for _o in (analysis.get('stat_obs') or []):
        if _o['ok']:
            _out.append('観察・統計紐(記録のみ) %s %s %d-%d  %s  紐統計%.1f%% 軸統計%.1f%%  (7/16〜9/13 %d点 的中%.1f%% 回収%.1f%%)'
                        % (_o['tag'], _jp[_o['kind']], _o['axis_ban'], _o['ban'], _o['label'],
                           _o['p_pt'], _o['p_ax'], _o['n'], _o['hit'], _o['roi']))
    return _out


def write_stat_obs_log(html_races, path):
    """★v136_020: 1レース1行の記録CSV(10月以降の答え合わせ用)。統計予想が1レースも無ければ書かない。"""
    import csv as _csv
    if not any(a.get('stat') for _, a, _ in html_races):
        return None
    _jp = {'umatan': '馬単', 'umaren': '馬連', 'wide': 'ワイド'}
    _cols = ['日付', '競馬場', 'R', '統計HTML']
    _cols += ['勝負レース', '単指数', '紐の並び', '参考買い目']
    for _d in STAT_OBS_DEFS:
        _cols += [f"{_d['tag']}({_d['axis']}{_jp[_d['kind']]})", f"{_d['tag']}紐統計%"]
    _cols += ['horselist結合', '騎手変更', '出走取消', '統計印', '版']   # ★v136_021
    _rows = []
    for name, a, _ in html_races:
        _st = a.get('stat')
        _date = ''
        if _st and _st.get('race_key'):
            _date = str(_st['race_key'][0])
        elif STAT_TARGET_DATE:
            _date = STAT_TARGET_DATE
        try:
            _rno = int(unicodedata.normalize('NFKC', str(name[2])).replace('R', '').strip())
        except Exception:
            _rno = name[2]
        _r = {'日付': _date, '競馬場': unicodedata.normalize('NFKC', str(name[0])).strip(), 'R': _rno,
              '統計HTML': 'あり' if _st else 'なし'}
        _r['勝負レース'] = '1' if a.get('shobu') else '0'
        _r['単指数'] = ('%.1f' % float(a.get('shobu_idx'))) if a.get('shobu_idx') is not None else ''
        _r['紐の並び'] = {'win': '単勝確率順', 'ev': '期待値順', 'legacy': '妙味重視(従来)'}.get(a.get('partner_mode'), '')
        _r['参考買い目'] = ' / '.join('%s %s' % (_kn, ','.join(_combos)) for _kn, _combos, _lab, _mks in _ref_bet_lines(a))
        _obs = {o['tag']: o for o in (a.get('stat_obs') or [])}
        for _d in STAT_OBS_DEFS:
            _o = _obs.get(_d['tag'])
            _key = f"{_d['tag']}({_d['axis']}{_jp[_d['kind']]})"
            _r[_key] = ('%d-%d' % (_o['axis_ban'], _o['ban'])) if (_o and _o['ok']) else ''
            _r[f"{_d['tag']}紐統計%"] = ('%.1f' % _o['p_pt']) if _o else ''
        try:
            _tb = a.get('table')
            _r['統計印'] = ' '.join('%s%d' % (str(_m), int(_b)) for _b, _m in zip(_tb['番'], _tb['統計印'])
                                   if str(_m or '').strip())
        except Exception:
            _r['統計印'] = ''
        _r['horselist結合'] = '%d/%d' % (a.get('hl_m') or 0, a.get('hl_n') or 0)
        _r['騎手変更'] = ' / '.join('%s番 %s→%s' % (c['ban'], c['old'], c['new']) for c in (a.get('dq_jk') or []))
        _r['出走取消'] = ' '.join('%d番%s' % (b, n) for b, n in sorted((a.get('dq_scratched') or {}).items()))
        _r['版'] = BADO_VERSION
        _rows.append(_r)
    with open(path, 'w', encoding='utf-8-sig', newline='') as _fh:
        _w = _csv.DictWriter(_fh, fieldnames=_cols)
        _w.writeheader()
        _w.writerows(_rows)
    return path


def _stat_block_html(analysis):
    _st = analysis.get('stat')
    if not _st:
        if STAT_DEBA_RACES:
            return '<div class="srow snone">統計予想: このレースの出馬表HTMLが見つからないため表示なし</div>'
        return ''
    _tb = analysis.get('table')
    _nm = {}
    try:
        _nm = {int(r['番']): str(r['馬名']) for _, r in _tb.iterrows()}
    except Exception:
        pass
    _mk = ' '.join('%s%d %s(統計%.1f%%)' % (_st['marks'][b], b, _h(_nm.get(b, '')), _st['p_stat'][b])
                   for b in _st['order'][:4])
    _jp = {'umaren': '馬連', 'wide': 'ワイド'}
    _bt = ' / '.join('%s %d-%d(%.1f%%)' % (_jp[x['kind']], x['ban1'], x['ban2'], x['prob'])
                     for x in _st['bets'])
    _extra = ''
    _ura = analysis.get('stat_ura') or {}
    if _ura:
        _seen = {}
        for _k in STAT_AXIS_CODES:
            if _k in _ura:
                _seen.setdefault(_ura[_k]['ban'], []).append(_k)
        _ul = ' / '.join('%d番 %s(統計%.1f%%)%s' % (
            _b, '・'.join(_ks), _ura[_ks[0]]['ps'],
            ' <b class="urayes">%s</b>' % STAT_URA_LABEL if _ura[_ks[0]]['ok'] else ' 裏付けなし')
            for _b, _ks in _seen.items())
        _extra += '<br><span class="sbt">軸の統計勝率(50%%以上=%s): %s</span>' % (STAT_URA_LABEL, _ul)
    _obs = [o for o in (analysis.get('stat_obs') or []) if o['ok']]
    if _obs:
        _extra += ('<br><span class="sbt">観察・統計紐（記録のみ）: %s</span>' % ' / '.join(
            '%s %s %d-%d(紐統計%.1f%%)' % (o['tag'], _jp[o['kind']], o['axis_ban'], o['ban'], o['p_pt'])
            for o in _obs))
    return ('<div class="srow"><span class="stag">統計予想（統計勝率のみ・参考・実弾外）</span>'
            '<span class="smk">%s</span><br><span class="sbt">参考買い目 %s</span>%s</div>' % (_mk, _bt, _extra))
REF_N_PARTNERS = 3             # 参考買い目の点数(=○▲△の頭数)
# 妙味の重み λ。★v137_001: 4.0→1.0→2.0。λ=4は人気薄に寄りすぎ(9/29 59R試算で○▲△の7割が
#   8番人気以下)、λ=1は9/27-28の2日68R集計で妙味を効かせなさすぎた(λ=2の方が馬連ワイドとも
#   回収率が上回った)。9/27-28の的中実績で比較検証済み。日を追えて再検証すること。
REF_VALUE_LAMBDA = _gz_env_num('REF_VALUE_LAMBDA', 2.0, float)
REF_PARTNER_MARKS = ('○', '▲', '△')
# 本線枠/次点本線枠が既に買っている組を参考から除くか。
#   False(既定): 参考買い目は常に ◎-○/▲/△ の3点(印と完全連動)。
REF_EXCLUDE_MAIN = False


def ref_partner_order(win_pct, adj_odds, axis_pos, lam=None):
    """◎(axis_pos 番目)の相手候補を、スコア降順の位置リストで返す(◎自身は含まない)。"""
    lam = REF_VALUE_LAMBDA if lam is None else float(lam)
    pm, pk, p = _judge_blend(win_pct, adj_odds)
    i = int(axis_pos)
    n = len(p)
    if n <= 1:
        return []
    with np.errstate(divide='ignore', invalid='ignore'):
        qb = p[i] * p / (1.0 - p[i]) + p * p[i] / np.clip(1.0 - p, 1e-12, None)
        qk = pk[i] * pk / (1.0 - pk[i]) + pk * pk[i] / np.clip(1.0 - pk, 1e-12, None)
        score = np.log(qb + 1e-12) + lam * np.log((qb + 1e-12) / (qk + 1e-12))
    score = np.where(np.isfinite(score), score, -1e9)
    score[i] = -np.inf
    return [int(k) for k in np.argsort(-score, kind='mergesort') if int(k) != i]


def compute_judge_estimates(win_pct, adj_odds, axis_pos):
    """◎(axis_pos 番目)の推定勝率・推定3着内率・対抗との馬連確率を返す。
    win_pct : 予想表の単勝確率(%, レース内合計≒100)
    adj_odds: 予想表の単勝オッズ(騎手補正後 adjusted_odds)。0以下/欠損は欠損扱い。
    戻り値 dict(est_win, est_top3, quinella, rival_win, rival_pos) すべて % 表記。"""
    pm, pk, p = _judge_blend(win_pct, adj_odds)
    i = int(axis_pos)
    est_win = float(p[i])
    h3 = min(max(_judge_harville_top3(p, i), 1e-4), 1 - 1e-4)
    z = JUDGE_TOP3_SLOPE * np.log(h3 / (1 - h3)) + JUDGE_TOP3_INTERCEPT
    est_top3 = float(1.0 / (1.0 + np.exp(-z)))
    order = [int(k) for k in np.argsort(-p, kind='mergesort') if int(k) != i]
    if order:
        j = order[0]
        q = p[i] * p[j] / (1 - p[i]) + p[j] * p[i] / (1 - p[j])
        rival_win = float(p[j])
    else:
        j, q, rival_win = None, 0.0, 0.0
    return dict(est_win=est_win * 100.0, est_top3=est_top3 * 100.0,
                quinella=float(q) * 100.0, rival_win=rival_win * 100.0, rival_pos=j)


def judge_grade_letter(est_win):
    for g, cut in JUDGE_GRADE_CUTS:
        if est_win >= cut:
            return g
    return 'D'


def judge_race_cat(est_win, est_top3, quinella, rival_win):
    """(race_cat, recommended_action) を返す。ラベルは従来と同一。"""
    if est_win >= JUDGE_CAT_INVEST_WIN and est_top3 >= JUDGE_CAT_INVEST_TOP3:
        return '投資レース', '軸の信頼度が特に高いレース(推定3着内率%g%%以上)。回収率100%%超の保証はない' % JUDGE_CAT_INVEST_TOP3
    if est_top3 >= JUDGE_CAT_STRONG_TOP3:
        return '有力レース', '軸が堅い。軸から馬連・ワイド中心'
    if JUDGE_CAT_TWO_ENABLE and quinella >= JUDGE_CAT_TWO_Q and rival_win >= JUDGE_CAT_TWO_RIVAL:
        return '有力レース（2頭軸）', '上位2頭軸で連系（馬連・ワイド）中心。単勝は見送り'
    if est_top3 < JUDGE_CAT_SKIP_TOP3:
        return '見送りレース', '基本的に見送り推奨'
    return '標準レース', 'ワイド中心に薄く'


def calculate_confidence_score(analysis):
    """
    v130: 的中信頼度(0-100)。「このレースの推奨買い目が的中しやすいか」を表す。
      核=軸の複勝確率/勝率(較正済み確率、的中率と単調対応で汎化)。補正=支配度(z)とNCS級。
      confidence_formula.json(v2:確率ブレンド)があればバックテスト最適化(target=hit)の
      重み・較正を採用。無ければ埋め込み既定(複勝主導ブレンド)で妥当に動作。
    """
    place = float(analysis.get('axis_place_prob', 0.0) or 0.0)
    win   = float(analysis.get('axis_win_prob', 0.0) or 0.0)
    z     = float(analysis.get('axis_zscore', 0.0) or 0.0)
    judgment_class = str(analysis.get('judgment_class', ''))
    for g in ('S', 'A', 'B', 'C'):
        if f'{g}級軸' in judgment_class:
            grade = g; break
    else:
        grade = 'D'

    est_hit = None
    used_src = '埋め込み既定'
    conf = None
    # ★v136_008: 新判定 = ◎の推定3着内率(%)をそのまま信頼度にする
    _new_top3 = analysis.get('judge_est_top3') if USE_NEW_JUDGE else None
    if _new_top3 is not None:
        conf = float(_new_top3)
        used_src = 'v136_009 推定3着内率'
    elif _CONF_FORMULA and 'blend' in _CONF_FORMULA:
        try:
            conf, est_hit = _conf_apply_formula(analysis, _CONF_FORMULA, place, win, z, grade)
            if conf is not None:
                used_src = _CONF_FORMULA.get('_src', 'confidence_formula.json')
        except Exception:
            conf = None
    if conf is None:
        D = _CONF_DEFAULT
        base = D['w_place'] * place + D['w_win'] * win
        z_adj = float(np.clip(z * D['z_gain'], D['z_clip'][0], D['z_clip'][1]))
        conf = base + z_adj + D['grade_adj'].get(grade, 0.0)
    score = int(round(min(100.0, max(0.0, conf))))

    if _new_top3 is not None:
        for _mn, _lb, _fh, _bg, _bar in JUDGE_CONF_TIERS:
            if score >= _mn:
                tier_label, font_hex, bg_hex, bar_hex = _lb, _fh, _bg, _bar
                break
    else:
        tier_cuts = (_CONF_FORMULA or {}).get('tier_cutoffs') if (_CONF_FORMULA and 'blend' in _CONF_FORMULA) else None
        tier_label, font_hex, bg_hex, bar_hex = _conf_tier(score, tier_cuts)

    n_umaren = len(analysis.get('umaren_recs', []))
    n_umatan = len(analysis.get('umatan_recs', []))
    n_wide = len(analysis.get('wide_recs', []))   # ★v136_001
    tan_tier = analysis.get('nsl_tan_tier')
    has_bets = (tan_tier in ('本命', '主軸', '中オッズ', '妙味')) or (n_umaren > 0) or (n_umatan > 0)

    best_ev = 0.0
    for key in ('tan_recs', 'fuku_recs', 'umaren_recs', 'umatan_recs', 'wide_recs'):
        for rec in analysis.get(key, []):
            try:
                best_ev = max(best_ev, float(rec.ev))
            except Exception:
                pass

    est_txt = f' 推定的中率~{est_hit:.0f}%' if est_hit is not None else ''
    breakdown = (
        f"軸 複勝率~{place:.0f}% 勝率~{win:.0f}% z{z:+.2f} {grade}級"
        f" → 信頼度{score}/100 [{tier_label}]{est_txt}"
    )
    if not has_bets:
        breakdown += '  ※推奨買い目なし'

    return {
        'score': score, 'category': tier_label, 'confidence': score, 'tier': tier_label,
        'est_hit': est_hit, 'font_hex': font_hex, 'bg_hex': bg_hex, 'bar_hex': bar_hex,
        'breakdown': breakdown, 'best_ev': int(best_ev), 'has_bets': has_bets, 'source': used_src,
    }

def _safe_positive(v, missing):
    """正の数値だけを返す。非数値/NaN/0以下は missing を返す共通実装。"""
    try:
        n = float(v)
    except Exception:
        return missing
    return missing if (n != n or n <= 0) else n


def safe_pop(v, missing=float('nan')):
    """人気(順位)を安全に取得。0 / 負 / 非数値は『欠損』として NaN を返す。

    【バグ修正】CSV読込時に 単勝人気 の欠損を 0 で埋めているため(numeric_cols の
    fillna(0))、素直に比較すると 0 が『1番人気より上』とみなされ、
    「人気5番以内」「人気2番以内」「人気3番以内」等の条件を全部すり抜けて
    実在しない買い目・安定運用判定を生んでいた。NaN にしておけば
    NaN <= 5 が False となり、欠損馬は全ての人気条件から自動的に外れる。"""
    return _safe_positive(v, missing)


def safe_odds(v, missing=float('nan')):
    """単勝オッズを安全に取得。0 / 負 / 非数値は『欠損』として NaN を返す。
    (0埋めされた欠損オッズが『オッズ<10』条件に該当してしまうのを防ぐ)"""
    return _safe_positive(v, missing)


import re as _re
import math as _math

# ════════════════════════════════════════════════════════════════
# ★ v135_017: レース分類「見送りレース」の表示のみ一旦中止
# ────────────────────────────────────────────────────────────────
#   race_cat の判定ロジック自体はそのまま残し、表示(judgment_class /
#   recommended_action)に出るときだけ中立表記へ差し替える。
#   v135_016では投資/見送りの両方を伏せていたが、投資レースは復活。
#   見送りレースも再開するときは _RACE_CAT_MASK から該当行を消すか、
#   SHOW_RACE_CAT_ALL = True に戻すだけでよい。
# ════════════════════════════════════════════════════════════════
SHOW_RACE_CAT_ALL = False   # False: _RACE_CAT_MASK のレース分類を表示しない
_RACE_CAT_MASK = {
    # '投資レース' は v135_017 で表示復活(マスク対象から除外)
    '見送りレース': ('標準レース', 'ワイド中心に薄く'),
}


def _mask_race_cat(race_cat, recommended_action):
    """表示用に race_cat / recommended_action を差し替える。"""
    if SHOW_RACE_CAT_ALL:
        return race_cat, recommended_action
    if race_cat in _RACE_CAT_MASK:
        return _RACE_CAT_MASK[race_cat]
    return race_cat, recommended_action

NSL_SKIP_VENUES     = set()  # 見送り開催地スキップなし(全開催で買い目を出す)。
NSL_REFERENCE_ONLY_VENUES = set()   # 参考強制する開催地: なし
NSL_TIER_REFERENCE_ONLY = set()     # 参考強制する信頼度区分: なし


# ══════════════════════════════════════════════════════════════════
# ★v137_001: 勝負レース(本線枠・次点本線枠は廃止)
#   ◎の単指数が90以上(=単勝確率 TAN_IDX_IRON_WP% 以上)のレースを『勝負レース』として表示する。
#   紐(○▲△)は、勝負レースは単勝確率順の上位3頭、それ以外のレースは期待値(単勝確率×予想オッズ)順の上位3頭。
#   (2026/09/17〜29の506Rで、単指数90以上は単勝確率順、それ以外は期待値順が回収率で有利だった。
#    REF_PARTNER_MODE=legacy で従来の ref_partner_order(妙味重視 λ=2)に戻せる)
#   実弾の買い目(本線枠・紐カバー枠)は出さない。買い目は参考予想(◎→○▲△ 馬連・ワイド・馬単各3点)として表示する。
# ══════════════════════════════════════════════════════════════════
GZ_COND_DEFS = {}
GZ_MAIN_KEEP = set()
REF_UMAREN_ENABLE = False

GZ_TIER_LABEL = {'main': '本線枠', 'cover': '紐カバー枠'}
GZ_KIND_JP = {'umatan': '馬単', 'umaren': '馬連', 'wide': 'ワイド'}
GZ_KIND_SEP = {'umatan': '→', 'umaren': '-', 'wide': '-'}
# 参考予想(◎→○▲△)の券種
REF_HIMO_BET_TYPES = ('umaren', 'wide', 'umatan')   # 馬連・ワイド・馬単(◎→紐)
# 勝負レース以外で表示しない券種(ワイドは勝負レース以外の回収率が低いため非表示)
REF_HIDE_KINDS_NON_SHOBU = ('wide',)
JITEN_ENABLE = False
JITEN_MAP_DEFS = []
_AXIS_LABEL = {}
SHOBU_LABEL = '勝負レース'
REF_PARTNER_MODE = str(_os_gz.environ.get('REF_PARTNER_MODE', 'auto')).strip().lower()   # auto / legacy


def gz_tier(cname):
    return GZ_COND_DEFS.get(cname, {}).get('tier', 'cover')


def _gz_names(kind):
    return [k for k, v in sorted(GZ_COND_DEFS.items(), key=lambda kv: kv[1]['order']) if v['kind'] == kind]


GZ_UMATAN_NAMES = _gz_names('umatan')
GZ_UMAREN_NAMES = _gz_names('umaren')
GZ_WIDE_NAMES = _gz_names('wide')


def partner_order_auto(win_pct, adj_odds, bans, axis_pos, shobu):
    """◎(axis_pos)の相手候補を並べた位置リスト。勝負レース=単勝確率順 / それ以外=期待値順(単勝確率×予想オッズ)。"""
    win = np.asarray(win_pct, dtype=float)
    od = np.clip(np.asarray(adj_odds, dtype=float), 1.0, None)
    idx = [i for i in range(len(win)) if i != axis_pos]
    if shobu:
        return sorted(idx, key=lambda i: (-win[i], bans[i]))
    return sorted(idx, key=lambda i: (-(win[i] * od[i]), bans[i]))




# ══════════════════════════════════════════════════════════════════
# ★v137_001: 確率の計算(オッズ勝率・人気オッズ・アンサンブル・着順確率)
# ══════════════════════════════════════════════════════════════════
def _normalize(v):
    v = np.clip(np.asarray(v, dtype=float), 0.0, None)
    s = v.sum()
    return v / s if s > 0 else np.full(len(v), 1.0 / max(len(v), 1))


def popularity_probs(group):
    """人気オッズの確率(条件付きロジット)。group の f_* 列を使う。無い列は 0。"""
    z = np.zeros(len(group))
    for k, b in POP_COEF.items():
        col = 'f_' + k
        if col in group.columns:
            z = z + b * pd.to_numeric(group[col], errors='coerce').fillna(0.0).to_numpy(dtype=float)
    z = z - z.max() if len(z) else z
    return _normalize(np.exp(z))


def odds_win_probs(raw_odds, p_fill):
    """補正前単勝オッズ → オッズ勝率(合計1)。オッズの無い馬は p_fill(人気オッズの確率)で埋める。"""
    od = np.asarray(pd.to_numeric(pd.Series(raw_odds), errors='coerce'), dtype=float)
    has = np.isfinite(od) & (od > 0)
    if not has.any():
        return _normalize(p_fill), has
    w = np.where(has, np.power(1.0 / np.where(has, od, 1.0), ODDS_PROB_POWER), 0.0)
    w = w / w.sum()
    if not has.all():
        # オッズの無い馬: 人気オッズの確率を、オッズのある馬の人気オッズ合計に対する比で差し込む
        pf = np.asarray(p_fill, dtype=float)
        scale = 1.0 / max(pf[has].sum(), 1e-9)
        w = np.where(has, w, pf * scale)
    return _normalize(w), has


def ensemble_win_probs(p_odds, p_stat):
    """単勝確率 ∝ オッズ勝率^ENS_W_ODDS × 統計勝率^ENS_W_STAT(p_stat が None ならオッズ勝率^ENS_ODDS_ONLY_POWER)。"""
    p_odds = _normalize(p_odds)
    if p_stat is None:
        return _normalize(np.power(np.clip(p_odds, 1e-9, None), ENS_ODDS_ONLY_POWER))
    p_stat = _normalize(p_stat)
    lz = (ENS_W_ODDS * np.log(np.clip(p_odds, 1e-6, None))
          + ENS_W_STAT * np.log(np.clip(p_stat, 1e-6, None)))
    return _normalize(np.exp(lz - lz.max()))


TAN_IDX_IRON_WP = _gz_env_num('TAN_IDX_IRON_WP', 62.5, float)   # 単勝確率(%)がこの値で単指数=90(鉄板クラス)
TAN_IDX_IRON = 90.0


def _axis_is_iron(analysis):
    try:
        t = analysis['table']
        return bool((t.loc[t['印'] == '◎', '単指数'] >= TAN_IDX_IRON).any())
    except Exception:
        return False


def tan_index(p_win):
    """単指数(100点満点)。単勝確率p(0〜1)を 100*(1-10^(-p/p90)) で換算(p90=TAN_IDX_IRON_WP/100)。
    既定は単勝確率62.5%で90点=鉄板クラス(0.805/0.345の確率。旧確率の55%と同じ約8%のレースが該当し、
    2026/09/17〜29の506Rで◎の単勝的中は約78%・複勝は全的中)。
    30%→ 約67点 / 45%→ 約81点 / 80%→ 約95点。単調増加なので単指数の大小は単勝確率の大小と同じ。"""
    p = np.clip(np.asarray(p_win, dtype=float), 0.0, None)
    return 100.0 * (1.0 - np.power(10.0, -p / (TAN_IDX_IRON_WP / 100.0)))


def finish_probs(p, lam2=None, lam3=None):
    """勝率 p(合計1)から Harville 型で着順確率を出す。2着・3着は p^λ で割り引く。
    戻り値 dict: top2/top3(各馬), exacta[i,j](i1着j2着), quinella[i,j](馬連), wide[i,j](ワイド)。"""
    lam2 = PLACE_LAMBDA2 if lam2 is None else lam2
    lam3 = PLACE_LAMBDA3 if lam3 is None else lam3
    p = _normalize(p)
    n = len(p)
    top2 = p.copy()
    top3 = np.zeros(n)
    ex = np.zeros((n, n))
    wide = np.zeros((n, n))
    if n < 2:
        return dict(top2=np.ones(n), top3=np.ones(n), exacta=ex, quinella=ex.copy(), wide=wide)
    a2 = np.power(p, lam2)
    a3 = np.power(p, lam3)
    s2, s3 = a2.sum(), a3.sum()
    for i in range(n):
        d2 = s2 - a2[i]
        if d2 <= 1e-12:
            continue
        for j in range(n):
            if j == i:
                continue
            pij = p[i] * a2[j] / d2
            ex[i, j] = pij
            top2[j] += pij
            if n < 3:
                top3[i] += pij
                top3[j] += pij
                wide[i, j] += pij
                continue
            d3 = s3 - a3[i] - a3[j]
            if d3 <= 1e-12:
                continue
            pk = pij * a3 / d3
            pk[i] = 0.0
            pk[j] = 0.0
            tot = pk.sum()
            top3[i] += tot
            top3[j] += tot
            top3 += pk
            wide[i, j] += tot           # i-j が1・2着
            wide[i, :] += pk            # i-k(1・3着)
            wide[j, :] += pk            # j-k(2・3着)
    q = ex + ex.T
    wide = wide + wide.T
    np.fill_diagonal(wide, 0.0)
    return dict(top2=np.clip(top2, 0, 1), top3=np.clip(top3, 0, 1), exacta=ex, quinella=q, wide=np.clip(wide, 0, 1))


def _empty_analysis():
    return {
        'table': pd.DataFrame(), 'axis_analysis': 0, 'axis_class': 'D級軸（混戦・軸不適）',
        'payout_level_score': 0, 'arare_class': '', 'race_trend': '', 'special_single': 'なし',
        'fuku_recs': [], 'umatan_recs': [], 'umaren_recs': [], 'wide_recs': [], 'tan_recs': [],
        'hole_fuku_recs': [], 'reference_bets': [], 'jiten_recs': [],
        'gz_umatan_cond': {}, 'gz_umaren_cond': {}, 'gz_wide_cond': {}, 'gz_stakes': {}, 'shobu': False, 'shobu_idx': 0.0, 'partner_mode': '',
        'gz_active_conds': [], 'force_reference': True, 'judgment_class': 'D級軸 / 標準レース',
        'recommended_action': '', 'expected_win_rate': '—', 'expected_place_rate': '—',
        'judgment_comment': 'データなし', 'myomi_axis_score': 0, 'axis_zscore': 0.0,
        'axis_win_prob': 0.0, 'axis_place_prob': 0.0, 'hole_bans': [], 'fav_bans': [],
    }


def _stat_probs_for_group(group, p_odds):
    """出馬表HTML × bado_stat_model の統計勝率(%)を {馬番: %} で返す。結合できなければ (None, None)。"""
    if not (STAT_ENABLE and STAT_DEBA_RACES and _bsm is not None):
        return None, None
    _m = stat_model()
    if not _m:
        return None, None
    try:
        _venue = re.sub(r'\s', '', unicodedata.normalize('NFKC', str(group['場所'].iloc[0])))
        _rno = _race_no_of(group['レース'].iloc[0])
        _names = {int(b): str(n) for b, n in zip(group['番'], group['馬名'])}
    except Exception:
        return None, None
    _race = _stat_find_race(_venue, _rno, _names)
    if _race is None:
        return None, None
    _race2 = _bsm.DebaRace(_race.date, _race.venue, _race.rno, _race.post,
                           [h for h in _race.horses if h['uma'] in _names], _race.path)
    if len(_race2.horses) < 2:
        return None, None
    _pex = {int(b): float(p) * 100.0 for b, p in zip(group['番'], p_odds)}
    try:
        _res = _bsm.predict_race(_m, _race2, p_exist=_pex)
    except Exception as _e:
        print(f'  [統計予想] {_venue}{_rno}R の計算に失敗: {_e}')
        return None, None
    return {int(b): float(r['p_stat']) * 100.0 for b, r in _res.items()}, _race


def generate_analysis(group):
    """★v137_001: 1レース分の予想(オッズ×統計のアンサンブル)。戻り値の dict は v136_021 と同じキー。"""
    if len(group) == 0:
        return _empty_analysis()
    group = group.reset_index(drop=True)
    group['番'] = group['番'].map(_to_int_z2h)
    group = group[group['番'].notna()].reset_index(drop=True)
    group['番'] = group['番'].astype(int)

    # ── 出走取消/除外(出馬表HTML) ──
    _dq_scratched = {}
    try:
        _dq_scratched = stat_scratched_for_group(group)
        if _dq_scratched and len(group) - len(_dq_scratched) >= 2:
            group = group[~group['番'].isin(set(_dq_scratched))].reset_index(drop=True)
        else:
            _dq_scratched = {}
    except Exception as _e_scr:
        print(f'  [出走取消] 判定に失敗: {_e_scr}')
        _dq_scratched = {}
    if len(group) < 2:
        return _empty_analysis()
    try:
        _dq_label = '%s%sR' % (re.sub(r'\s', '', unicodedata.normalize('NFKC', str(group['場所'].iloc[0]))),
                               _race_no_of(group['レース'].iloc[0]))
    except Exception:
        _dq_label = ''
    if '騎手指数' not in group.columns:
        group['騎手指数'] = 50.0
    group['騎手指数'] = pd.to_numeric(group['騎手指数'], errors='coerce').fillna(50.0)
    _dq_jk = [dict(ban=int(r['番']), name=str(r['馬名']), old=str(r['騎手_旧']),
                   new=str(r['騎手']).replace(JK_CHANGE_MARK, ''), idx_old=None, idx_new=float(r['騎手指数']))
              for _, r in group.iterrows() if str(r.get('騎手_旧', '') or '').strip()]
    _dq_hl_n = len(group)
    _dq_hl_m = int(pd.Series(group['hl']).astype(bool).sum()) if 'hl' in group.columns else 0

    n = len(group)
    raw_odds = pd.to_numeric(group['単勝オッズ'], errors='coerce')
    # ── 人気オッズ・オッズ勝率 ──
    p_pop = popularity_probs(group)
    p_odds, _has_odds = odds_win_probs(raw_odds, p_pop)
    pop_odds = np.clip(TAKEOUT_WIN / np.clip(p_pop, 1e-6, None), 1.0, 999.9)
    # ── 予想オッズ(補正前オッズ × 人気オッズ、補正前と同じ控除率にそろえる) ──
    ro = raw_odds.to_numpy(dtype=float)
    lo = np.where(_has_odds, POP_ODDS_W_RAW * np.log(np.where(_has_odds, ro, 1.0))
                  + (1 - POP_ODDS_W_RAW) * np.log(pop_odds), np.log(pop_odds))
    od = np.exp(lo)
    _target = (np.sum(1.0 / ro[_has_odds]) / max(_has_odds.mean(), 1e-9)) if _has_odds.any() else 1.0 / TAKEOUT_WIN
    od = od * (np.sum(1.0 / od) / _target)
    group['前売オッズ'] = raw_odds.round(1)
    group['人気オッズ'] = np.round(pop_odds, 1)
    group['adjusted_odds'] = np.round(np.clip(od, 1.0, 999.9), 1)
    group['オッズ勝率'] = np.round(p_odds * 100.0, 1)

    # ── 統計勝率 → 単勝確率(アンサンブル) ──
    try:
        _ps_map, _st_race = _stat_probs_for_group(group, p_odds)
    except Exception as _e_st:
        print(f'  [統計予想] 失敗: {_e_st}')
        _ps_map, _st_race = None, None
    if _ps_map:
        _ps = np.array([_ps_map.get(int(b), np.nan) for b in group['番']], dtype=float) / 100.0
        _ps = np.where(np.isfinite(_ps) & (_ps > 0), _ps, p_odds * np.nansum(_ps))   # 欠けた馬はオッズ勝率で埋める
        p_stat = _normalize(_ps)
        group['統計勝率'] = np.round(p_stat * 100.0, 1)
    else:
        p_stat = None
        group['統計勝率'] = np.nan
    p_win = ensemble_win_probs(p_odds, p_stat)
    fp = finish_probs(p_win)
    group['単勝確率'] = np.round(p_win * 100.0, 1)
    group['複勝確率'] = np.round(fp['top3'] * 100.0, 1)
    group['単指数'] = np.round(tan_index(p_win), 1)                    # 単勝確率の100点満点換算(90以上=鉄板クラス)
    group['単勝期待値'] = np.round(p_win * group['adjusted_odds'].to_numpy(dtype=float) * 100.0).astype(int)
    # ── 紐馬指数: 2〜3着に来る確率 × 予想オッズ^HIMO_ODDS_POWER (レース内最大=100) ──
    _sub = np.clip(fp['top3'] - p_win, 1e-6, None)
    _hs = _sub * np.power(np.clip(group['adjusted_odds'].to_numpy(dtype=float), 1.0, None), HIMO_ODDS_POWER)
    group['紐馬指数'] = np.round(100.0 * _hs / _hs.max(), 1)
    # ── 人気(予想オッズ順) / 前売人気 ──
    _ord = sorted(range(n), key=lambda i: (group.at[i, 'adjusted_odds'], group.at[i, '番']))
    _pop = np.zeros(n, dtype=int)
    for r_, i in enumerate(_ord, start=1):
        _pop[i] = r_
    group['adjusted_popularity'] = _pop
    _rp = pd.to_numeric(group.get('単勝人気'), errors='coerce') if '単勝人気' in group.columns else None
    if _rp is None or _rp.isna().all():
        _rp = raw_odds.rank(method='min')
    group['前売人気'] = _rp
    group[KISHU_DEV_COL] = group['騎手指数'].round(1)
    if '展開' not in group.columns:
        group['展開'] = '-'
    group['印'] = ''

    idx_of_ban = {int(b): i for i, b in enumerate(group['番'])}

    def _v(i, col):
        try:
            x = float(group.at[i, col])
            return x if x == x else None
        except Exception:
            return None

    def _r1(i, col):
        x = _v(i, col)
        return None if x is None else float('%.1f' % x)

    def _ge(i, col, thr):
        x = _r1(i, col)
        return x is not None and x >= thr

    def _odds_lt(i, lim):
        x = _v(i, 'adjusted_odds')
        return x is not None and 0 < x < lim

    def _pop_of(i):
        return int(group.at[i, 'adjusted_popularity'])

    _wp_order = sorted(range(n), key=lambda i: (-p_win[i], group.at[i, '番']))

    # ── 予想印: ◎=連対確率1位 / ○▲△=勝負レース(◎の単指数90以上)は単勝確率順、それ以外は期待値順 ──
    axis_idx = min(range(n), key=lambda i: (-fp['top2'][i], group.at[i, '番']))
    group.at[axis_idx, '印'] = '◎'
    _shobu = float(group.at[axis_idx, '単指数']) >= TAN_IDX_IRON
    _partners = []
    try:
        if REF_PARTNER_MODE == 'legacy':
            _ordp = ref_partner_order(group['単勝確率'].to_numpy(), raw_odds.to_numpy(), axis_idx)
        else:
            _ordp = partner_order_auto(group['単勝確率'].to_numpy(), group['adjusted_odds'].to_numpy(dtype=float),
                                       group['番'].to_numpy(), axis_idx, _shobu)
        _partners = [k for k in _ordp[:REF_N_PARTNERS]]
    except Exception as _re_:
        print(f'  [参考予想警告] 相手選定に失敗: {_re_}')
        _partners = [i for i in _wp_order if i != axis_idx][:REF_N_PARTNERS]
    for _mk, i in zip(REF_PARTNER_MARKS, _partners):
        group.at[i, '印'] = _mk
    second_idx = _partners[0] if _partners else None

    # ── 穴: 前売人気4〜9位で 単勝確率がオッズ勝率の1.3倍以上・複勝確率25%以上 ──
    anaba_hidx_bans = []
    for i in range(n):
        _p = _v(i, '前売人気')
        if _p is None or not (ANA_POP_RANGE[0] <= _p <= ANA_POP_RANGE[1]):
            continue
        if p_win[i] >= ANA_RATIO_MIN * p_odds[i] and fp['top3'][i] * 100.0 >= ANA_PLACE_MIN:
            anaba_hidx_bans.append(int(group.at[i, '番']))
            _cur = str(group.at[i, '印'] or '')
            if _cur == '':
                group.at[i, '印'] = '穴'
            elif _cur != '◎':
                group.at[i, '印'] = _cur + '穴'

    # ── 買い目: 本線枠・次点本線枠は廃止。◎→○▲△(馬連・ワイド・馬単各3点)を参考予想として表示する ──
    #   勝負レース(◎の単指数90以上)は紐が単勝確率順、それ以外は期待値順。実弾の買い目は出さない。
    umatan_recs, umaren_recs, wide_recs = [], [], []
    gz_umatan_cond, gz_umaren_cond, gz_wide_cond = {}, {}, {}
    jiten_recs = []
    _ax_ban = int(group.at[axis_idx, '番'])
    if REF_PARTNER_MODE == 'legacy':
        _ref_desc = '◎→○▲△(配当の妙味を重視した上位3頭)'
    elif _shobu:
        _ref_desc = '◎→○▲△(勝負レース: 単勝確率の高い上位3頭)'
    else:
        _ref_desc = '◎→○▲△(期待値の高い上位3頭)'
    reference_bets = []
    if _partners:
        for kd in REF_HIMO_BET_TYPES:
            if not _shobu and kd in REF_HIDE_KINDS_NON_SHOBU:
                continue
            reference_bets.append(dict(kind=kd, name='', axis_ban=_ax_ban,
                                       partner_bans=[int(group.at[i, '番']) for i in _partners],
                                       stats={}, is_fallback=False, shobu=_shobu,
                                       desc=_ref_desc))

    # ── 統計予想(参考) ──
    _stat, _stat_ura = None, {}
    if p_stat is not None:
        _ps_d = {int(b): float(v) for b, v in zip(group['番'], group['統計勝率'])}
        _order = sorted(_ps_d, key=lambda b: (-_ps_d[b], b))
        _marks = {b: STAT_MARKS[k] for k, b in enumerate(_order[:len(STAT_MARKS)])}
        _bets = []
        if len(_order) >= 3:
            _fs = finish_probs(p_stat)
            _ia = idx_of_ban[_order[0]]
            for kd, tab in (('umaren', _fs['quinella']), ('wide', _fs['wide'])):
                for b in _order[1:3]:
                    _bets.append(dict(kind=kd, ban1=_order[0], ban2=b, prob=round(tab[_ia, idx_of_ban[b]] * 100, 1)))
        _stat = dict(race_key=getattr(_st_race, 'key', None), path=getattr(_st_race, 'path', ''),
                     p_stat=_ps_d, p_fused={int(b): float(v) for b, v in zip(group['番'], group['単勝確率'])},
                     marks=_marks, order=_order, bets=_bets)
        group['統計印'] = group['番'].map(lambda b: _marks.get(int(b), ''))
    else:
        group['統計印'] = ''

    # ── 表示テーブル(単勝確率降順) ──
    table_columns = ['番', '馬名', '騎手', KISHU_DEV_COL, '展開', '前売オッズ', '人気オッズ', '前売人気',
                     '単指数', '紐馬指数', '印', 'adjusted_popularity', 'adjusted_odds',
                     '単勝確率', '複勝確率', '単勝期待値', '統計勝率', 'オッズ勝率', '統計印']
    table_df = group[table_columns].copy()
    table_df['_o'] = -p_win
    table_df = table_df.sort_values(['_o', '番']).drop(columns=['_o'])
    table_df = table_df.rename(columns={'adjusted_popularity': '単勝人気', 'adjusted_odds': '単勝オッズ'})

    # ── 軸級・信頼度・レース分類(◎の単勝確率・複勝確率) ──
    a_win = float(p_win[axis_idx] * 100.0)
    a_top3 = float(fp['top3'][axis_idx] * 100.0)
    _rv = next((i for i in _wp_order if i != axis_idx), None)
    judge_est = dict(est_win=a_win, est_top3=a_top3,
                     quinella=(float(fp['quinella'][axis_idx, _rv] * 100.0) if _rv is not None else 0.0),
                     rival_win=(float(p_win[_rv] * 100.0) if _rv is not None else 0.0))
    _g = judge_grade_letter(a_win)
    race_cat, recommended_action = judge_race_cat(a_win, a_top3, judge_est['quinella'], judge_est['rival_win'])
    race_cat, recommended_action = _mask_race_cat(race_cat, recommended_action)
    _grade_desc = {'S': 'S級軸（1強・高信頼）', 'A': 'A級軸（有力）', 'B': 'B級軸（標準）',
                   'C': 'C級軸（やや不安）', 'D': 'D級軸（混戦・軸不適）'}
    axis_class = _grade_desc[_g]
    judgment_class = '%s / %s' % (axis_class, race_cat)
    _wp_sorted = sorted(p_win * 100.0, reverse=True)
    a_gap = round(a_win - (_wp_sorted[1] if n > 1 else 0.0), 1)
    myomi_axis_score = int(round(float(np.clip(a_gap * 3.0 + (a_win - 20.0) * 1.5, 0, 90))))
    _sd = float(np.std(p_win * 100.0))
    axis_zscore = (a_win - 100.0 / n) / _sd if _sd > 0 else 0.0
    axis_pop = _pop_of(axis_idx)
    payout_level_score = round(float(np.clip(20 + axis_pop * 5 + max(0.0, 8.0 - _sd) * 2.0, 5, 100)), 1)
    arare_class = ('低配当予想（本命固め）' if payout_level_score <= 45 else
                   '中配当予想（やや荒れ）' if payout_level_score <= 60 else '高配当予想（大荒れ）')

    # ── コメント ──
    a_row = group.loc[axis_idx]
    a_odds = float(a_row['adjusted_odds'])
    a_ev = int(a_row['単勝期待値'])
    _ev_word = ('非常に高い(妙味大)' if a_ev >= 130 else '高い(妙味あり)' if a_ev >= 110 else
                '標準的(ほぼ適正オッズ)' if a_ev >= 95 else '低い(過剰人気ぎみ)')
    expected_win_rate = '単勝期待値%d(%s)' % (a_ev, _ev_word)
    expected_place_rate = '勝率%.0f%%/複勝率%.0f%%・予想オッズ%.1f倍' % (a_win, a_top3, a_odds)
    _kj = float(a_row['騎手指数'])
    _k_word = '好騎手' if _kj >= 60 else ('標準的な騎手' if _kj >= 45 else 'やや割引く騎手')
    _st_txt = ('統計%.1f%%' % float(a_row['統計勝率'])) if p_stat is not None else '統計なし'
    judgment_comment = ' / '.join(x for x in [
        '◎%d番 %s' % (_ax_ban, a_row['馬名']),
        ('鞍上%s(騎手指数%.0f・%s)' % (a_row['騎手'], _kj, _k_word)) if a_row['騎手'] else '',
        'オッズ勝率%.1f%%・%s' % (float(a_row['オッズ勝率']), _st_txt),
        '%d番人気(前売%s倍)' % (axis_pop, ('%.1f' % a_row['前売オッズ']) if a_row['前売オッズ'] == a_row['前売オッズ'] else '-'),
        '単勝確率%.1f%%・複勝確率%.1f%%' % (a_win, a_top3)] if x)
    race_trend = '◎[%d] 単勝確率%.1f%%(2位差%.1f) ＋ 相手%s' % (
        _ax_ban, a_win, a_gap, [int(group.at[i, '番']) for i in _wp_order if i != axis_idx][:4])

    # ── 賭け金(均等買い・本線枠の点数上限) ──
    _has_conn = bool(umatan_recs or umaren_recs or wide_recs)
    _venue = str(group['場所'].iloc[0]).strip()
    _conf_tier_label = ''
    try:
        _conf_tier_label = calculate_confidence_score({'judge_est_top3': a_top3}).get('category', '')
    except Exception:
        pass
    _tier_ref = _conf_tier_label in NSL_TIER_REFERENCE_ONLY
    force_reference = (_venue in NSL_REFERENCE_ONLY_VENUES) or _tier_ref or (not _has_conn)
    _gz_max_pts = gz_race_max_points()
    _all = sorted([('umatan', (r.ban1, r.ban2), p) for p, r in enumerate(umatan_recs)]
                  + [('umaren', (r.ban1, r.ban2), p) for p, r in enumerate(umaren_recs)]
                  + [('wide', (r.ban1, r.ban2), p) for p, r in enumerate(wide_recs)],
                  key=lambda t: (GZ_COND_DEFS[{'umatan': gz_umatan_cond, 'umaren': gz_umaren_cond,
                                               'wide': gz_wide_cond}[t[0]][t[1]]]['order'], t[2]))
    _keep = set((k, key) for k, key, _p in _all[:_gz_max_pts])
    gz_dropped_n = max(0, len(_all) - _gz_max_pts)
    if gz_dropped_n:
        umatan_recs = [r for r in umatan_recs if ('umatan', (r.ban1, r.ban2)) in _keep]
        umaren_recs = [r for r in umaren_recs if ('umaren', (r.ban1, r.ban2)) in _keep]
        wide_recs = [r for r in wide_recs if ('wide', (r.ban1, r.ban2)) in _keep]
        gz_umatan_cond = {k: v for k, v in gz_umatan_cond.items() if ('umatan', k) in _keep}
        gz_umaren_cond = {k: v for k, v in gz_umaren_cond.items() if ('umaren', k) in _keep}
        gz_wide_cond = {k: v for k, v in gz_wide_cond.items() if ('wide', k) in _keep}
    _amt = 0 if force_reference else GZ_FLAT_UNIT
    gz_stakes = {}
    for _kind, _cm in (('umatan', gz_umatan_cond), ('umaren', gz_umaren_cond), ('wide', gz_wide_cond)):
        for k in _cm:
            gz_stakes[(k[0], k[1], _kind)] = _amt
    _gz_active = sorted(set(gz_umatan_cond.values()) | set(gz_umaren_cond.values()) | set(gz_wide_cond.values()),
                        key=lambda c: GZ_COND_DEFS.get(c, {}).get('order', 99))
    yuuryoku_umatan = {(r.ban1, r.ban2) for r in umatan_recs}
    yuuryoku_wide = {(r.ban1, r.ban2) for r in wide_recs}

    _t_idx = _wp_order[0]
    tan_recs = [SingleRec(ban=int(group.at[_t_idx, '番']), hit_rate=round(float(p_win[_t_idx] * 100.0), 1),
                          odds=float(group.at[_t_idx, 'adjusted_odds']), ev=int(group.at[_t_idx, '単勝期待値']))]
    fav_bans = [int(b) for b, p in zip(group['番'], group['adjusted_popularity']) if 1 <= p <= 3]
    axis_analysis = round(float(np.clip(a_win, 0, 100)), 1)

    return {
        'table': table_df, 'axis_analysis': axis_analysis, 'axis_class': axis_class,
        'payout_level_score': payout_level_score, 'arare_class': arare_class, 'race_trend': race_trend,
        'special_single': '◎%d 単勝確率%.1f%% 複勝確率%.0f%%' % (_ax_ban, a_win, a_top3),
        'fuku_recs': [], 'umatan_recs': umatan_recs, 'umaren_recs': umaren_recs, 'wide_recs': wide_recs,
        'tan_recs': tan_recs, 'trio_recs': [], 'yuryoku_umatan_pairs': [],
        'anaba_umaren': [], 'anaba_umatan': [], 'anaba_wide': [], 'is_invest': False,
        'yuuryoku_tan': set(), 'yuuryoku_fuku': set(), 'yuuryoku_umatan': yuuryoku_umatan,
        'yuuryoku_umaren': set(), 'yuuryoku_wide': yuuryoku_wide,
        'myomi_umaren': set(), 'myomi_umatan': set(), 'force_reference': force_reference,
        'stable_tan_rec': None, 'stable_tan_tag': '', 'stable_fuku_rec': None, 'stable_fuku_tag': '',
        'subline_rec': None, 'subline_b_rec': None, 'subline_c_rec': None,
        'jiten_recs': jiten_recs, 'stat': _stat, 'stat_ura': _stat_ura, 'stat_obs': [],
        'shobu': bool(_shobu), 'shobu_idx': float(group.at[axis_idx, '単指数']),
        'partner_mode': ('legacy' if REF_PARTNER_MODE == 'legacy' else ('win' if _shobu else 'ev')),
        'hl_n': _dq_hl_n, 'hl_m': _dq_hl_m, 'dq_jk': _dq_jk, 'dq_scratched': _dq_scratched,
        'dq_nostat': (p_stat is None), 'race_label': _dq_label,
        'tan_priority_bans': [], 'conf_tier_precheck': _conf_tier_label,
        'hole_fuku_recs': [], 'high_confidence': False, 'conf_score': 0.0, 'maruren_1pt': None,
        'wide_1pt': None, 'popular_out_high_payout': False, 'pop_conf_score': 0.0,
        'pop_maruren_recs': [], 'pop_wide_recs': [], 'myomi_axis_score': myomi_axis_score,
        'hole_bans': [], 'fav_bans': fav_bans, 'is_hole_axis': False, 'is_honmei_axis': False,
        'axis_ban': _ax_ban, 'sub_axis_ban': (int(group.at[second_idx, '番']) if second_idx is not None else None),
        'axis_score': a_win, 'himo_bans': [int(group.at[i, '番']) for i in _partners],
        'axis_tan_roi': 0.0, 'axis_tan_n': 0, 'axis_fuku_roi': 0.0, 'axis_tan_cond': '',
        'judgment_class': judgment_class, 'recommended_action': recommended_action,
        'expected_win_rate': expected_win_rate, 'expected_place_rate': expected_place_rate,
        'judgment_comment': judgment_comment, 'judgment_score': round(float(np.clip(0.7 * axis_analysis + 0.3 * myomi_axis_score, 0, 100)), 1),
        'axis_zscore': round(float(axis_zscore), 2),
        'axis_win_prob': a_win, 'axis_place_prob': a_top3,
        'judge_est_win': round(a_win, 2), 'judge_est_top3': round(a_top3, 2),
        'judge_quinella': round(judge_est['quinella'], 2),
        'nsl_tan_tier': '', 'nsl_fuku_hit': False, 'spec_axis_type': '%s級軸' % _g,
        'tan_flags': {}, 'fuku_flags': {}, 'gap_conf_factor': 1.0, 'two_axis_mode': False,
        'reference_bets': reference_bets,
        'gz_umatan_cond': gz_umatan_cond, 'gz_umaren_cond': gz_umaren_cond, 'gz_wide_cond': gz_wide_cond,
        'anaba_hidx_bans': anaba_hidx_bans, 'gz_stakes': gz_stakes,
        'gz_active_conds': _gz_active, 'gz_bankroll': GZ_BANKROLL,
        'gz_flat_unit': GZ_FLAT_UNIT, 'gz_max_points': _gz_max_pts,
        'gz_dropped_n': gz_dropped_n, 'gz_race_invest': _amt * len(gz_stakes),
        'gz_active_main': list(_gz_active), 'gz_active_ana': [],
        'gz_main_points': len(gz_stakes), 'gz_ana_points': 0,
        'gz_main_invest': _amt * len(gz_stakes), 'gz_ana_invest': 0, 'gz_ana_max_points': 0,
    }


# ══════════════════════════════════════════════════════════════════
# ★v137_001: 入力の読込 ── 単勝オッズCSV(nar_shutuba_all) × horselist × 出馬表HTML
#   read_racecard_csv(path) が1日分の出馬表 DataFrame を返す(GUI・一括予想の共通入口)。
#   列: 場所 距離 レース 出走時刻 番 枠 馬名 性齢 斤量 騎手 厩舎 単勝オッズ(補正前) 単勝人気(補正前)
#       日付 展開 hl(horselist結合) 騎手_旧(騎手変更時) 騎手指数 f_*(人気オッズの材料)
# ══════════════════════════════════════════════════════════════════
# netkeiba の地方競馬 場コード(レースID の 5〜6 桁目)
NAR_VENUE_CODE = {
    '30': '門別', '31': '北見', '32': '岩見沢', '33': '帯広', '34': '旭川', '35': '盛岡', '36': '水沢',
    '37': '上山', '38': '三条', '39': '足利', '40': '宇都宮', '41': '高崎', '42': '浦和', '43': '船橋',
    '44': '大井', '45': '川崎', '46': '金沢', '47': '笠松', '48': '名古屋', '49': '紀三井寺', '50': '園田',
    '51': '姫路', '52': '益田', '53': '福山', '54': '高知', '55': '佐賀', '56': '荒尾', '57': '中津',
    '58': '札幌', '59': '函館', '60': '新潟', '61': '中京', '65': '帯広',
}
_ODDS_COL_ALIASES = {
    '番': ('馬番', '馬 番', '番'), '単勝オッズ': ('オッズ', '単勝オッズ', '単勝'),
    '単勝人気': ('人気', '単勝人気'), '馬名': ('馬名',), '騎手': ('騎手',), '枠': ('枠', '枠番'),
    '性齢': ('性齢',), '斤量': ('斤量',), '厩舎': ('厩舎',), 'レースID': ('レースID', 'race_id'),
    'レース情報': ('レース情報',),
}


def _rec4(x):
    """'a-b-c-d' → [a, b, c, d]。読めなければ [0,0,0,0]。"""
    try:
        v = [int(t) for t in unicodedata.normalize('NFKC', str(x)).split('-')]
        return v if len(v) == 4 else [0, 0, 0, 0]
    except Exception:
        return [0, 0, 0, 0]


def _eb_logit(num, n, prior, k=3.0):
    """経験ベイズで縮約した率の logit。num/n が少数でも暴れないよう prior へ k 走分寄せる。"""
    p = (num + k * prior) / (n + k)
    p = min(max(p, 1e-4), 1 - 1e-4)
    return float(np.log(p / (1 - p)))


def _tsec(x):
    """'1:03.3' → 63.3 秒。"""
    m = re.match(r'^\s*(\d+):(\d+(?:\.\d+)?)', unicodedata.normalize('NFKC', str(x or '')))
    return (int(m.group(1)) * 60 + float(m.group(2))) if m else np.nan


def _name_key(x):
    return re.sub(r'[\s\u3000]', '', unicodedata.normalize('NFKC', str(x or '')))


def _dist_txt(d):
    try:
        return '%dm' % int(float(d))
    except Exception:
        return '距離不明'


def _deba_meta(race):
    """出馬表HTML(DebaRace)から (発走時刻, 距離m) を取る。取れなければ ('', '')。"""
    post, dist = '', ''
    if race is None:
        return post, dist
    try:
        _p = str(getattr(race, 'post', '') or '').strip()
        _m = re.search(r'(\d{1,2})[:：時](\d{2})', unicodedata.normalize('NFKC', _p))
        if _m:
            post = '%02d:%s' % (int(_m.group(1)), _m.group(2))
    except Exception:
        pass
    for _a in ('dist', 'distance', 'kyori', '距離'):
        _v = getattr(race, _a, None)
        if _v:
            _n = _to_int_z2h(re.sub(r'[^0-9]', '', unicodedata.normalize('NFKC', str(_v))))
            if _n and 300 <= _n <= 4000:
                return post, str(_n)
    try:
        _path = getattr(race, 'path', None)
        if _path and _os.path.exists(str(_path)):
            _raw = open(str(_path), 'rb').read()
            _txt = None
            for _enc in ('utf-8', 'cp932', 'euc_jp'):
                try:
                    _txt = _raw.decode(_enc)
                    break
                except Exception:
                    continue
            if _txt:
                _txt = unicodedata.normalize('NFKC', re.sub(r'<[^>]+>', ' ', _txt)).replace(',', '')
                _m = (re.search(r'(?:ダート|ダ|芝)\s*(\d{3,4})\s*m', _txt)
                      or re.search(r'(?<!\d)(\d{3,4})\s*m(?![a-z])', _txt))
                if _m and 300 <= int(_m.group(1)) <= 4000:
                    dist = _m.group(1)
                if not post:
                    _m = re.search(r'発走(?:時刻)?\s*[:：]?\s*(\d{1,2})[:：時](\d{2})', _txt)
                    if _m:
                        post = '%02d:%s' % (int(_m.group(1)), _m.group(2))
    except Exception:
        pass
    return post, dist


_LEG_ORDER = ('逃', '先', '差', '追')
# 出馬表HTML『母馬名』の行(馬主名と並べて、各過去走のタイム・コーナー通過順・上がり3Fが
# 入っている特殊なレイアウト。実物で確認済み)。例: "1:00.6　7-9　37.2"(タイム/コーナー通過順/上がり3F)。
_PAST_TIME_RE = re.compile(r'(\d{1,2}):(\d{2})\.(\d)[\u3000\s]+([\d\-]+)[\u3000\s]+(\d{1,2})\.(\d)')
_HORSE_NUM_RE = re.compile(r'class="horseNum">([^<]*)</td>')
_FIELD_SIZE_RE = re.compile(r'(\d{1,2})頭')          # raceInfo(着順+日付+馬場+頭数)の頭数


def _corner_positions(raw):
    """'7-9' / '1-1-4-4' → [7,9] / [1,1,4,4](コーナー通過順=スタートからの位置取り)。"""
    return [int(x) for x in re.split(r'-', raw) if x.strip().isdigit()]


def _leg_style_of_run(field, pos, last3f):
    """1走分の脚質を推定。pos=コーナー通過順(先頭が最初の通過点)、field=その走の出走頭数。
    平均通過順位(÷頭数、0=先頭・1=最後方)だけで判定する。
    ★境界値は 2026/09/29 の実データ(5場60レース・639頭)で調整した。当初は
    道中の進出量(最初の通過順−最後の通過順)で差し/追込を分けていたが、地方競馬は
    コーナー通過が2〜3回しか記録されない(直線競馬・短距離)レースが多く、その進出量の
    中央値がほぼ0(ratio_avg>0.45の走の中央値0.0、四分位0.0〜0.11)で差がほとんど
    付かず、大半が追込側に倒れて『追込ばかりになる』偏りが出ていた(実測52.0%)。
    平均通過順位だけの単純な帯分けに変えたところ 逃13%/先25%/差35%/追27% と
    大きな偏りなく分かれたため、この形に統一する。last3f は現状未使用(タイム・頭数の
    異なる過去走間で単純比較できないため)だが、将来の精緻化用にデータは保持する。"""
    if not pos or not field or field < 2:
        return None
    ratio_avg = (sum(pos) / len(pos)) / field
    if ratio_avg <= 0.18:
        return '逃'
    if ratio_avg <= 0.45:
        return '先'
    if ratio_avg <= 0.75:
        return '差'
    return '追'


def _leg_style_from_runs(runs):
    """runs = [(頭数, コーナー通過順list, 上がり3F), ...](先頭が最新走)。
    直近最大3走ぶんの脚質を多数決で決める(同数は最新走を優先)。判定できる過去走が
    無ければ空文字を返す(表示は'-')。"""
    labels = [lb for lb in (_leg_style_of_run(*r) for r in runs[:3]) if lb]
    if not labels:
        return ''
    counts = {lb: labels.count(lb) for lb in _LEG_ORDER}
    top = max(counts.values())
    cands = {lb for lb in _LEG_ORDER if counts[lb] == top}
    return next(lb for lb in labels if lb in cands)    # labels は新しい順 → 最新走を優先


def _deba_leg_styles(race):
    """出馬表HTML(DebaRace)から {馬番: 脚質(逃/先/差/追)} を、各馬の直近走の
    タイム・コーナー通過順(スタートからの位置取り)・上がり3Fタイムから算出して返す
    (出馬表HTMLに脚質そのものの記載は無いため、この3値から推定する)。
    raw HTML を『class="horseNum">N</td>』の出現位置で馬ごとのブロックに切り、
    ブロック内の raceInfo(頭数)と タイム・コーナー通過順・上がり3F(同じ並び順)を
    対応づける。取得・算出に失敗してもレースを落とさない(空dictを返す)。"""
    out = {}
    if race is None:
        return out
    try:
        _path = getattr(race, 'path', None)
        if not (_path and _os.path.exists(str(_path))):
            return out
        _raw = open(str(_path), 'rb').read()
        _txt = None
        for _enc in ('utf-8', 'cp932', 'euc_jp'):
            try:
                _txt = _raw.decode(_enc)
                break
            except Exception:
                continue
        if not _txt:
            return out
        _idxs = [(m.start(), m.group(1)) for m in _HORSE_NUM_RE.finditer(_txt)]
        _idxs = [(pos, n) for pos, n in _idxs if n.isdigit()]      # 見出し行(『馬番』)を除く
        for _k, (pos, numstr) in enumerate(_idxs):
            end = _idxs[_k + 1][0] if _k + 1 < len(_idxs) else len(_txt)
            block = _txt[pos:end]
            fields = [int(x) for x in _FIELD_SIZE_RE.findall(block)]   # 各過去走の出走頭数(前走→)
            runs = []
            for m in _PAST_TIME_RE.finditer(block):
                corner = _corner_positions(m.group(4))
                if not corner:
                    continue
                fsz = fields[len(runs)] if len(runs) < len(fields) else None
                if fsz:
                    runs.append((fsz, corner, float('%s.%s' % (m.group(5), m.group(6)))))
            lg = _leg_style_from_runs(runs)
            if lg:
                out[int(numstr)] = lg
    except Exception:
        pass
    return out


def _deba_has_html(folder):
    import glob as _g
    return bool(_g.glob(_os.path.join(str(folder), '**', 'R[0-9]*_*.htm*'), recursive=True))


def _deba_autoload(folder):
    """出馬表HTML(R??_*.html)を、オッズCSVのフォルダ → 予想コード本体のフォルダ の順に
    (サブフォルダも含めて)探して読む。対象日のレースが見つからなければ次の場所も試す。"""
    if STAT_DEBA_RACES or _bsm is None:
        return
    _code_dir = _os.path.dirname(_os.path.abspath(__file__))
    _cands = []
    for _d in (folder, _code_dir):
        if _d and str(_d) not in _cands and _os.path.isdir(str(_d)):
            _cands.append(str(_d))
    for _d in _cands:
        try:
            if not _deba_has_html(_d):
                continue
            stat_load_deba_folder(_d)
            if not STAT_TARGET_DATE or any(k[0] == STAT_TARGET_DATE for k in STAT_DEBA_RACES):
                return
            print(f'  [統計予想] {_d} に対象日({STAT_TARGET_DATE})の出馬表がないため次の場所も探します。')
        except Exception as _e:
            print(f'  [統計予想] 出馬表HTMLの自動読込に失敗({_d}): {_e}')


def _read_odds_csv(path):
    df = _read_csv_auto(path)
    _ren = {}
    _cols = {re.sub(r'\s+', '', c): c for c in df.columns}
    for canon, aliases in _ODDS_COL_ALIASES.items():
        for a in aliases:
            a2 = re.sub(r'\s+', '', a)
            if a2 in _cols and _cols[a2] not in _ren:
                _ren[_cols[a2]] = canon
                break
    df = df.rename(columns=_ren)
    df = df.loc[:, ~df.columns.duplicated()]
    if '枠' in df.columns:                           # 2行目に見出しが繰り返されている
        df = df[df['枠'].astype(str).str.strip() != '枠']
    need = ['番', '馬名', '騎手', '単勝オッズ', 'レースID']
    miss = [c for c in need if c not in df.columns]
    if miss:
        raise ValueError(f'単勝オッズCSV(nar_shutuba_all)に必要な列がありません: {miss}')
    return df.reset_index(drop=True)


def is_odds_csv(path):
    """単勝オッズCSV(レースID と オッズ の列がある)かどうかを先頭行で判定。"""
    try:
        with open(path, 'rb') as fh:
            first = fh.read(4096).split(b'\n')[0]
        for enc in ('utf-8-sig', 'cp932'):
            try:
                t = first.decode(enc)
            except Exception:
                continue
            if 'レースID' in t and 'オッズ' in t:
                return True
    except Exception:
        pass
    return False


def read_racecard_csv(path, require_cols=None):
    """★v137_001: 単勝オッズCSV + horselist + 出馬表HTML から1日分の出馬表を作る。"""
    global CURRENT_RACECARD_DIR, CURRENT_RACECARD_DATE, _HL_MISSING_WARNED
    CURRENT_RACECARD_DIR = _os.path.dirname(_os.path.abspath(str(path)))
    CURRENT_RACECARD_DATE = _guess_date_from_name(path)
    _HL_MISSING_WARNED = False
    try:
        stat_set_target_date_from_name(Path(str(path)).name)
    except Exception:
        pass
    _deba_autoload(CURRENT_RACECARD_DIR)

    src = _read_odds_csv(path)
    rows = []
    for r in src.to_dict('records'):
        rid = re.sub(r'\D', '', str(r.get('レースID') or ''))
        ban = _to_int_z2h(r.get('番'))
        if len(rid) < 12 or ban is None:
            continue
        venue = NAR_VENUE_CODE.get(rid[4:6], rid[4:6])
        d8 = int(rid[:4] + rid[6:10])
        rno = int(rid[10:12])
        od = pd.to_numeric(unicodedata.normalize('NFKC', str(r.get('単勝オッズ') or '')), errors='coerce')
        pp = pd.to_numeric(unicodedata.normalize('NFKC', str(r.get('単勝人気') or '')), errors='coerce')
        rows.append(dict(
            場所=venue, 日付=d8, レース=str(rno), 番=ban, 枠=_to_int_z2h(r.get('枠')),
            馬名=str(r.get('馬名') or '').strip(), 性齢=str(r.get('性齢') or '').strip(),
            斤量=pd.to_numeric(r.get('斤量'), errors='coerce'), 騎手=str(r.get('騎手') or '').strip(),
            厩舎=str(r.get('厩舎') or '').strip(),
            単勝オッズ=(float(od) if od == od and od > 0 else np.nan),
            単勝人気=(float(pp) if pp == pp and pp > 0 else np.nan),
            レース情報=str(r.get('レース情報') or ''), 展開='-'))
    if not rows:
        raise ValueError('単勝オッズCSVにレースID・馬番の揃った行がありません。')
    df = pd.DataFrame(rows)
    # ★v137_001: 帯広(ばんえい競走)は他場と競走形態が違う(重量そりを曳く輓曳競走)ため予想しない。
    _n_obihiro = int((df['場所'] == '帯広').sum())
    if _n_obihiro:
        df = df[df['場所'] != '帯広'].reset_index(drop=True)
        print(f'  [帯広除外] 帯広(ばんえい競走) {_n_obihiro}頭を予想対象から除外しました。')
    if df.empty:
        raise ValueError('帯広以外に予想できるレースがありません。')
    if CURRENT_RACECARD_DATE is None or CURRENT_RACECARD_DATE not in set(df['日付']):
        CURRENT_RACECARD_DATE = int(df['日付'].mode().iloc[0])
        stat_set_target_date(str(CURRENT_RACECARD_DATE))

    # ── horselist 結合(馬名の一致も確認) ──
    table = _get_horselist_table(CURRENT_RACECARD_DIR)
    hl = []
    for r in df.itertuples():
        h = table.get((_norm_venue_hl(r.場所), int(r.日付), int(r.レース), int(r.番)))
        if h is not None and _name_key(h.get('馬名')) != _name_key(r.馬名):
            h = None                                  # 馬番は合っても馬名が違う → 結合しない
        hl.append(h)
    df['hl'] = [h is not None for h in hl]

    # ── 騎手変更(horselist の騎手名を優先)。騎手キーは同じ場の騎手の正式名にそろえる ──
    df['騎手_旧'] = ''
    for i, h in enumerate(hl):
        if not h:
            continue
        new = re.sub(r'\s', '', unicodedata.normalize('NFKC', str(h.get('騎手名') or '')))
        old = df.at[i, '騎手']
        if new and new.lower() != 'nan' and old and not _jockey_same(old, new):
            df.at[i, '騎手_旧'] = old
            df.at[i, '騎手'] = new + JK_CHANGE_MARK
    _keys = {}
    for v, grp in df.groupby('場所'):
        names = sorted({_jk_nz(n) for n, o in zip(grp['騎手'], grp['騎手_旧']) if not o})
        for i in grp.index:
            k = _jk_nz(df.at[i, '騎手'])
            if df.at[i, '騎手_旧']:
                cand = [n for n in names if _jockey_same(n, k)]
                if len(cand) == 1:
                    k = cand[0]
            _keys[i] = (v, k)
    df['_jkey'] = pd.Series(_keys)

    # ── 人気オッズ・騎手指数の材料 ──
    df['_rk'] = df['場所'] + '|' + df['日付'].astype(str) + '|' + df['レース']
    inv = (1.0 / df['単勝オッズ'])
    pm = inv / inv.groupby(df['_rk']).transform('sum')
    nf = df['単勝オッズ'].notna().groupby(df['_rk']).transform('sum')
    df['_q'] = np.log(pm * nf)                                    # 市場の相対評価(0=平均)
    g = df.groupby('_jkey')['_q']
    tot, cnt = g.transform('sum'), g.transform('count')
    q0 = df['_q'].fillna(0.0)
    has_q = df['_q'].notna().astype(float)
    df['f_jloo'] = (tot - q0) / (cnt - has_q + JOCKEY_LOO_SHRINK)
    rides = df.groupby('_jkey')['番'].transform('size').astype(float)
    lr = np.log(rides)
    df['f_jride'] = lr - lr.groupby(df['場所']).transform('mean')

    def _hv(i, col):
        return (hl[i] or {}).get(col) if hl[i] else None
    f_app, f_lw, f_lv, f_lc, f_exp, t = [], [], [], [], [], []
    for i in range(len(df)):
        f_app.append(1.0 if re.search('[☆▲△◇★]', str(_hv(i, '負担重量') or '')) else 0.0)
        a = _rec4(_hv(i, '全成績'))
        f_lw.append(_eb_logit(a[0], sum(a), 0.09))
        f_exp.append(float(np.log1p(sum(a))))
        v = _rec4(_hv(i, '当競馬場成績'))
        f_lv.append(_eb_logit(sum(v[:3]), sum(v), 0.27))
        c = _rec4(_hv(i, '騎手成績'))
        f_lc.append(_eb_logit(sum(c[:3]), sum(c), 0.27))
        t.append(_tsec(_hv(i, '最高タイム')))
    df['f_app'], df['f_lw'], df['f_lv'], df['f_lc'], df['f_exp'] = f_app, f_lw, f_lv, f_lc, f_exp
    df['_t'] = t
    gt = df.groupby('_rk')['_t']
    df['f_bt'] = (-(df['_t'] - gt.transform('mean')) / gt.transform('std').replace(0, np.nan)).fillna(0.0)
    df['f_bt_na'] = df['_t'].isna().astype(float)

    # ── 騎手指数(0-100・50=その場の平均)。人気オッズの係数で合成して場ごとに偏差値化 ──
    lc_c = df['f_lc'] - df['f_lc'].groupby(df['場所']).transform('mean')
    jraw = (POP_COEF['jloo'] * df['f_jloo'] + POP_COEF['jride'] * df['f_jride']
            + POP_COEF['app'] * df['f_app'] + POP_COEF['lc'] * lc_c)
    gj = jraw.groupby(df['場所'])
    sd = gj.transform(lambda s: s.std(ddof=0)).replace(0, np.nan)
    df['騎手指数'] = (50.0 + 10.0 * (jraw - gj.transform('mean')) / sd).fillna(50.0).clip(0, 100).round(1)

    # ── 発走時刻・距離・脚質(出馬表HTML。レース情報列に書かれていればそれも使う) ──
    df['出走時刻'] = ''
    df['距離'] = ''
    n_leg = 0
    for rk, grp in df.groupby('_rk', sort=False):
        i0 = grp.index[0]
        post, dist = '', ''
        info = unicodedata.normalize('NFKC', str(df.at[i0, 'レース情報']))
        _m = re.search(r'(\d{3,4})\s*m', info)
        if _m:
            dist = _m.group(1)
        _m = re.search(r'(\d{1,2}):(\d{2})', info)
        if _m:
            post = '%02d:%s' % (int(_m.group(1)), _m.group(2))
        if STAT_DEBA_RACES and _bsm is not None:
            try:
                _race = _stat_find_race(re.sub(r'\s', '', df.at[i0, '場所']), int(df.at[i0, 'レース']),
                                        {int(b): str(n) for b, n in zip(grp['番'], grp['馬名'])})
                _p2, _d2 = _deba_meta(_race)
                post, dist = (post or _p2), (dist or _d2)
                _leg = _deba_leg_styles(_race)         # ★v137_001: {馬番: 逃/先/差/追}
                if _leg:
                    for i in grp.index:
                        _b = _to_int_z2h(df.at[i, '番'])
                        if _b in _leg:
                            df.at[i, '展開'] = _leg[_b]
                            n_leg += 1
            except Exception:
                pass
        df.loc[grp.index, '出走時刻'] = post
        df.loc[grp.index, '距離'] = dist
    n_hl = int(df['hl'].sum())
    print(f'[入力] {Path(str(path)).name}: {df["_rk"].nunique()}R {len(df)}頭 / horselist結合 {n_hl}/{len(df)} / '
          f'騎手変更 {int((df["騎手_旧"] != "").sum())} / 脚質判明 {n_leg}/{len(df)} / 出馬表HTML {stat_summary()}')
    return df.drop(columns=['_t', '_q', 'レース情報'])


# =====================================================
# ★ 新機能：一括予想（フォルダ指定） v079_一括予想版
# =====================================================
def _run_batch_prediction(app, var_status, var_out_dir, messagebox, filedialog, Path, pd):
    """フォルダ内の全CSVを連続処理して予想ファイルを出力（修正版: スレッド化 + 堅牢化）"""
    folder = filedialog.askdirectory(title='単勝オッズCSV(nar_shutuba_all_*.csv)が入ったフォルダを選択してください')
    if not folder:
        return

    # ★v137_001: レースID列を持つ単勝オッズCSVだけを対象にする(horselist・出力CSVは除く)
    csv_files = [p for p in sorted(Path(folder).glob("*.csv")) if is_odds_csv(p)]
    if not csv_files:
        messagebox.showwarning('CSVが見つかりません',
                               '選択したフォルダに単勝オッズCSV(レースID・オッズ列のある nar_shutuba_all_*.csv)がありません。')
        return

    # 出力先フォルダの解決（修正: 未設定・相対パス時の迷子を防止）
    #   - 空欄 / '.' / 相対パスの場合は、選択した入力フォルダ内の「予想結果」サブフォルダへ出力
    #   - 最終的に必ず絶対パスへ正規化し、完了ダイアログで実際の場所を明示する
    _out_raw = (var_out_dir.get() or '').strip()
    if (not _out_raw) or _out_raw in ('.', './') or (not Path(_out_raw).is_absolute()):
        out_dir = Path(folder) / '予想結果'
    else:
        out_dir = Path(_out_raw)
    out_dir = out_dir.resolve()
    out_dir.mkdir(parents=True, exist_ok=True)
    # 画面の出力先欄にも実際の絶対パスを反映しておく
    app.after(0, lambda p=str(out_dir): var_out_dir.set(p))

    total = len(csv_files)
    success_count = [0]  # mutable for closure
    error_files = []

    def batch_worker():
        for idx, csv_path in enumerate(csv_files, 1):
            app.after(0, lambda idx=idx, nm=csv_path.name: var_status.set(f'⏳ 一括処理中 ({idx}/{total}) : {nm}'))
            app.after(0, app.update_idletasks)

            try:
                # 1. データ読み込み（単一モードと同じロジックを再利用）
                #   ★ cp932/utf-8混在・列名ゆれ(予想 オッズ/人気)に対応
                df = read_racecard_csv(csv_path)

                df = _sort_by_race_no(df)
                grouped = df.groupby(_RACE_GROUP_KEYS, sort=False)

                if grouped.ngroups == 0:
                    # 空データでもファイルを作成してログを残す
                    stem = csv_path.stem
                    output_csv = out_dir / f"{stem}_予想結果_空データ_{datetime.now().strftime('%Y%m%d')}.csv"
                    with open(output_csv, 'w', encoding='utf-8-sig') as f:
                        f.write("有効なレースデータがありませんでした。\n")
                    raise ValueError("有効なレースデータがありません")

                # 2. 出力ファイル名生成（入力ファイル名ベース）
                stem = csv_path.stem
                import re
                m8 = re.search(r'(20\d{6})', stem)
                if m8:
                    date_str = m8.group(1)
                else:
                    m4 = re.search(r'(\d{4})', stem)
                    if m4 and len(m4.group(1)) == 4:
                        date_str = f"{datetime.now().year}{m4.group(1)}"
                    else:
                        date_str = datetime.now().strftime('%Y%m%d')

                # ★修正(重大): 旧版は出力名が日付のみに依存していたため、同一フォルダ内に
                #   同日付の複数CSVがあると後続ファイルが前のものを上書きし、
                #   「成功と表示されるのにファイルが1つしか残らない」問題が起きていた。
                #   入力ファイルの stem を含めて衝突を防ぐ。
                stat_set_target_date_from_name(stem)   # ★v136_021 GUIと同じ決め方(不明なら最新日+馬名照合)
                safe_stem = re.sub(r'[\\/:*?"<>|]', '_', stem)
                base_name = f"race_forecasts_{safe_stem}_{date_str}"
                output_csv = out_dir / f"{base_name}.csv"
                excel_path = out_dir / f"{base_name}_colored.xlsx"
                # 万一それでも衝突する場合は連番を付与
                _dup = 1
                while output_csv.exists() or excel_path.exists():
                    output_csv = out_dir / f"{base_name}_{_dup}.csv"
                    excel_path = out_dir / f"{base_name}_{_dup}_colored.xlsx"
                    _dup += 1

                # 3. 既存の出力関数を呼び出し
                _run_output(grouped, output_csv, excel_path)

                success_count[0] += 1

            except Exception as e:
                error_files.append(f"{csv_path.name}: {str(e)[:80]}")
                # ★修正: lambda がループ変数 csv_path を遅延参照すると最終値になるため束縛する
                app.after(0, lambda nm=csv_path.name: var_status.set(f'⚠️ エラー発生: {nm}'))
                app.after(0, app.update_idletasks)

        # 完了メッセージ（メインスレッドで表示）
        # 実際にファイルが生成されたか検証（「成功したのに無い」を検知）
        produced = sorted(out_dir.glob('race_forecasts_*.csv'))
        loc_line = f"出力先（絶対パス）:\n{out_dir}\n\n生成ファイル数: {len(produced)} 件"
        if error_files:
            msg = f"処理完了: {success_count[0]}/{total} 成功\n\n{loc_line}\n\nエラー発生ファイル:\n" + "\n".join(error_files[:5])
            if len(error_files) > 5:
                msg += f"\n...他 {len(error_files)-5} 件"
            app.after(0, lambda: messagebox.showwarning('一括処理完了（一部エラー）', msg))
        else:
            app.after(0, lambda: messagebox.showinfo('一括処理完了',
                f'✅ {total} ファイルの予想処理が正常に完了しました。\n\n{loc_line}'))

        app.after(0, lambda: var_status.set(f'✅ 一括予想完了（{success_count[0]}/{total} 成功）'))

    # ★ 修正: 別スレッドで実行（GUIフリーズ防止）
    threading.Thread(target=batch_worker, daemon=True).start()


# ══════════════════════════════════════════════════════════════
# ★ HTML予想表出力(中央競馬 v5ツールと同スタイル / NAR用)
#   新聞風: クリーム背景 + 深緑ヘッダ + レースカード + 枠色 + 印色。
#   買い目は採用条件を赤バナー+条件別色チップで強調表示。
# ══════════════════════════════════════════════════════════════
_NAR_WAKU_STYLE = {
    1: ("#ffffff", "#000"), 2: ("#222222", "#fff"), 3: ("#d43c3c", "#fff"),
    4: ("#2f6fd6", "#fff"), 5: ("#e6c822", "#000"), 6: ("#3c9e5f", "#fff"),
    7: ("#e08a3c", "#fff"), 8: ("#e88ab0", "#000"),
}
_NAR_MARK_COLOR = {'◎': '#c0392b', '○': '#1f6fd6', '▲': '#0e8f76',
                   '△': '#d4860b', '☆': '#6f42c1', '穴': '#b03a2e', '注': '#7d6608'}
_NAR_GZ_LABEL = {   # ★v135_003: 条件レジストリから自動生成 / ★v136_001: ワイド対応
    k: (v['label'].split(' ', 1)[0] + ' ' + GZ_KIND_JP.get(v['kind'], v['kind']),
        v['short'], v['bg'], v['fg'])
    for k, v in GZ_COND_DEFS.items()
}

_NAR_HTML_HEAD = """<!DOCTYPE html><html lang="ja"><head><meta charset="utf-8">
<title>地方競馬予想 (オッズ×統計 アンサンブル)</title><style>
 body{font-family:"Yu Gothic","Meiryo",sans-serif;margin:0;background:#f3f1ea;color:#1a1a1a}
 .top{background:#13392b;color:#fff;padding:10px 14px;position:sticky;top:0;z-index:5}
 .top h1{margin:0;font-size:18px;letter-spacing:1px}.top .sub{font-size:11px;opacity:.85;margin-top:3px}
 .legend{font-size:11px;margin-top:4px}.legend span{margin-right:8px}
 .race{background:#fff;margin:8px 10px;border:1px solid #cfc9bb;border-radius:6px;overflow:hidden;box-shadow:0 1px 3px rgba(0,0,0,.08)}
 .gz-race{border:3px solid #c00000;box-shadow:0 0 0 2px #ffe3e3 inset}
 .rhead{padding:4px 10px;background:#efece2;border-bottom:1px solid #d8d2c4;display:flex;align-items:center;gap:6px;flex-wrap:wrap}
 .rtitle{font-weight:bold;font-size:14px}.rmeta{font-size:11px;color:#555}
 .badge{margin-left:auto;font-weight:bold;font-size:12px;padding:2px 8px;border-radius:12px;color:#fff;background:#c00000}
 .rel{font-size:11px;font-weight:bold;padding:1px 7px;border-radius:10px;color:#fff}
 table.grid{width:100%;border-collapse:collapse;font-size:11px}
 .grid th{background:#13392b;color:#fff;padding:2px 3px;font-weight:normal;font-size:10px;white-space:nowrap}
 .grid td{padding:2px 3px;border-bottom:1px solid #eee;text-align:center;white-space:nowrap}
 .grid td.name{text-align:left;font-weight:bold}
 .grid td.jk{text-align:left;color:#444}
 .grid td.mk{font-size:14px;font-weight:bold}.grid td.waku{font-weight:bold;border-radius:3px}
 tr.axis{background:#fdf3e0}
 td.prob.tan{color:#c0392b}td.prob.fuku{color:#1f6fd6}
 td.hi-tan{background:#ffe0e0;color:#c0392b;font-weight:bold}      /* 単勝確率40%+ */
 td.hi-fuku{background:#dbeafe;color:#1743a0;font-weight:bold}     /* 複勝確率70%+ */
 td.hi-idx{background:#e5ffe5;color:#1e7a1e;font-weight:bold}      /* 単勝確率が前売オッズの評価より高い(妙味) */
 td.jk-hot{color:#c0392b;font-weight:bold}                        /* 騎手指数60+ の騎手名 */
 tr.anaba td{background:#fff3e0}                                   /* ★011 穴条件該当馬の行 */
 tr.anaba td.mk{color:#e67e22;font-weight:bold}
 /* ★v135_007: PDF(印刷)用。
    ・print-color-adjust:exact が無いと Chrome/Edge は背景色を落とすため、
      枠色・高単期待値の濃緑・穴行の色がPDFで全部消える。必須。
    ・表が横長なので landscape 固定。
    ・レースブロックと行が途中で分断されないようにする。 */
 @page{size:A4 landscape;margin:8mm}
 @media print{
   html,body{background:#fff;-webkit-print-color-adjust:exact;print-color-adjust:exact}
   *{-webkit-print-color-adjust:exact;print-color-adjust:exact}
   .race{page-break-inside:avoid;break-inside:avoid}
   .grid tr{page-break-inside:avoid;break-inside:avoid}
   .grid thead{display:table-header-group}
   .head{position:static}
 }
 /* ★v135_006: 以下2つは `tr.anaba td`(詳細度 class1+要素2)より強くする必要があるため
    `.grid` を付けて class2 にしている。順序も anaba より後に置く。 */
 .grid td.evhi{background:#1b6b3a;color:#ffffff;font-weight:bold}  /* 高単期待値(110+)=濃緑背景+白文字 */
 .grid td.kdev{color:#5b6b7a}                                      /* 補正騎手指数(表示専用) */
 .grid td.kdev.hi{color:#1f6f8b;font-weight:bold}                  /* 同 偏差値60+ */
 td.hi-ana{background:#ffe0b2;color:#a04a00;font-weight:bold}      /* 穴馬の人気オッズ */
 .gradebig{display:inline-block;font-weight:bold;font-size:15px;padding:2px 10px;border-radius:6px;letter-spacing:1px}
 .g-S{background:#ffd700;color:#000}.g-A{background:#ff7a45;color:#fff}
 .g-B{background:#58a6ff;color:#fff}.g-C{background:#3fb950;color:#fff}.g-D{background:#8b949e;color:#fff}
 .buys{padding:6px 10px;background:#faf8f2;border-top:1px dashed #cfc9bb}
 .gzbanner{background:#c00000;color:#fff;font-weight:bold;font-size:12px;padding:3px 8px;border-radius:4px;margin-bottom:4px}
 .gzactive{font-size:11px;color:#c00000;font-weight:bold;margin:1px 0 4px}
 .brow{margin:2px 0;display:flex;align-items:center;gap:4px;flex-wrap:wrap}
 .bt{display:inline-block;font-size:10.5px;font-weight:bold;color:#fff;background:#13392b;border-radius:7px;padding:1px 7px}
 .bt.anatier{background:#8a6d1f}
 .bt.subtier{background:#1f5f8a}
 .subnote{font-size:10px;color:#6b7280;margin-left:4px}
 .srow{margin:4px 10px 8px;padding:6px 9px;border:1px dashed #7b61a8;border-radius:6px;background:#f7f3fc;font-size:12px;line-height:1.7}
 .srow.snone{border-color:#c9c3d6;color:#8a8499;background:#faf9fc}
 .stag{display:inline-block;background:#5b3f8f;color:#fff;border-radius:5px;padding:0 7px;margin-right:6px;font-weight:bold;font-size:11px}
 .ura{display:inline-block;background:#5b3f8f;color:#fff;border-radius:4px;padding:0 5px;margin-left:4px;font-size:10px;font-weight:bold;vertical-align:middle}
 .urayes{color:#5b3f8f}
 .dnote{margin:4px 10px 6px;padding:5px 9px;border:1px solid #d4a017;border-radius:6px;background:#fff7e0;font-size:12px;color:#6b4e00}
 td.uracell{background:#e9dff7;font-weight:bold;color:#5b3f8f}
 .smk{font-weight:bold;color:#3f2a66}.sbt{color:#4b4260}
 td.sthi{background:#ece3fa;font-weight:bold;color:#3f2a66}
 .subwarn{color:#b45309;margin-left:4px}
 .bi{display:inline-block;background:#eee;border-radius:5px;padding:2px 7px;font-size:12px;margin:1px}
 .bi b{font-size:1.15em}.bi small{color:#666;margin-left:4px;font-size:10px}
 .stake{display:inline-block;font-weight:bold;color:#c00000;font-size:11.5px;margin-left:3px}
 .condchip{display:inline-block;font-size:10px;font-weight:bold;border-radius:7px;padding:1px 6px;margin-left:3px}
 .stbet{display:inline-flex;align-items:center;gap:6px;border-radius:7px;padding:3px 6px;margin:1px;border:2px solid transparent}
 .stbet.in{background:#eafaf0;border-color:#1b7a3a}
 .stbet.out{background:#f2f3f5;border-color:#c4c8cf}
 .stban{display:inline-flex;align-items:center;justify-content:center;min-width:24px;height:24px;background:#222;color:#fff;font-weight:bold;font-size:14px;border-radius:6px;padding:0 3px}
 .stnum{font-size:12px;color:#333}
 .stnum b{font-size:16px;color:#0a5c2b}
 .stbet.out .stnum b{color:#555}
 .stod{margin-left:6px;font-size:12px;color:#555;font-weight:bold}
 .sttag{display:inline-block;font-size:10.5px;font-weight:bold;color:#fff;border-radius:7px;padding:1px 7px;margin-left:2px}
 .skipnote{color:#888;font-size:11px}
 .skip{color:#999;font-size:12px;padding:3px}
 .footer{padding:10px;text-align:center;color:#777;font-size:10.5px}
</style></head><body>
<div class="top"><h1>BADO 地方競馬予想 — オッズ×統計 アンサンブル</h1>
<div class="sub">__SRC__ ／ 全__NR__レース ／ ★勝負レース(◎の単指数90以上) __NGZ__レース ／ 買い目は表示のみ</div>
<div class="legend">
__GZ_LEGEND__
 <span style="color:#ffcf9e">■穴=前売人気4-9位 × 単勝確率がオッズ勝率の1.21倍以上 × 複勝確率25%以上</span>
</div></div>
"""

NL_JOIN_GZ = '\n'


def _h(v):
    """HTMLエスケープ。★v135_010: 馬名/騎手名等に < > & " が入ると表が壊れる
    (XMLとしてパース不能になることを確認済み)。出力前に必ず通す。"""
    import html as _html
    return _html.escape('' if v is None else str(v), quote=True)


def _gz_legend_html():
    """★v137_001: HTMLレジェンド(勝負レース・確率とオッズ・統計予想・参考予想の説明)。"""
    _parts = []
    _parts.append(' <b>【%s】</b>' % SHOBU_LABEL)
    _parts.append(' <span style="color:#ffcf70">■◎の単指数が90以上(単勝確率%g%%以上)のレース。'
                  '紐(○▲△)は単勝確率の高い順に3頭。それ以外のレースは期待値(単勝確率×予想オッズ)の高い順に3頭。'
                  '買い目は◎→○▲△の馬連・ワイド・馬単各3点(表示のみ。勝負レース以外はワイドなし)。</span>' % TAN_IDX_IRON_WP)
    # ★v137_001: 確率・オッズの説明
    _parts.append(' <b>【確率とオッズ】</b>')
    _parts.append(' <span style="color:#d9c8f5">■単勝率=オッズ勝率(前売=補正前の単勝オッズを正規化)^%.2f × 統計%%^%.2f の融合。'
                  '複勝率・馬連/ワイドの的中率は単勝率から Harville 型(2着^%.2f・3着^%.2f で割引)。'
                  '単オッズ=予想オッズ(前売 %.0f%% × 人気O %.0f%% の対数平均、前売と同じ控除率)。'
                  '人気O=騎手・成績・当地・コンビ・最高タイムから推定した人気の単勝オッズ。'
                  '単指数=単勝確率の100点満点換算(90以上=鉄板クラス)、紐馬指数=2〜3着に来る確率×√予想オッズ(レース内最大100)。騎手指数=50が平均。</span>'
                  % (ENS_W_ODDS, ENS_W_STAT, PLACE_LAMBDA2, PLACE_LAMBDA3,
                     POP_ODDS_W_RAW * 100, (1 - POP_ODDS_W_RAW) * 100))
    if STAT_DEBA_RACES:
        _parts.append(' <b>【統計予想(参考・実弾外)】</b>')
        _parts.append(' <span style="color:#d9c8f5">■出馬表HTMLの近5走・着別成績・騎手/調教師成績から統計%%を計算(bado_stat_model)。'
                      '統計印◎○▲△=統計%%の上位。参考買い目=統計◎-○/◎-▲の馬連・ワイド(確率はHarville推定)。'
                      '読込: %s</span>' % _h(stat_summary()))
    else:
        _parts.append(' <b>【統計予想】</b> <span style="color:#ffb4a8">■出馬表HTML未読込 → 統計%なし・単勝率はオッズ勝率のみ</span>')
    _parts.append(' <b>【参考予想】</b>')
    _parts.append(' <span style="color:#9aa0a6">■紐 軸=◎(◎=連対率1位) × 予想印○▲△の3頭'
                  '(勝負レースは単勝確率順、それ以外は期待値順)。勝負レースは馬連・ワイド・馬単、それ以外は馬連・馬単を各3点表示。'
                  'いずれも表示のみで実弾対象外。</span>')
    return NL_JOIN_GZ.join(_parts)


# ══════════════════════════════════════════════════════════════════
# ★v135_007: HTML予想表を PDF にも出力する
#   要件: Windowsで追加インストールなしに動くことを最優先にしたいので、
#   バックエンドを以下の優先順で試し、最初に成功したものを使う。
#
#   1) Edge / Chrome のヘッドレス印刷 (--headless --print-to-pdf)
#      Windows10/11 は Edge が標準搭載なので追加インストール不要。
#      CSS(@page landscape / 背景色)をそのまま解釈するので再現性が最も高い。
#   2) Playwright (chromium) — pip install playwright + playwright install chromium
#   3) wkhtmltopdf — 外部バイナリがPATHにあれば使う
#   4) WeasyPrint — pip install weasyprint (Windowsは別途GTKが必要で難儀)
#
#   どれも無ければ PDF はスキップし、理由と導入方法を print する。
#   ★PDF化に失敗しても HTML/Excel/CSV の出力は絶対に止めない。
#
#   ※背景色(枠色・高単期待値の濃緑・穴行)は print-color-adjust:exact を
#     HTML側CSSに入れてあることが前提。無いとPDFで色が全部落ちる。
# ══════════════════════════════════════════════════════════════════
_PDF_CHROME_CANDIDATES = (
    r'C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe',
    r'C:\Program Files\Microsoft\Edge\Application\msedge.exe',
    r'C:\Program Files\Google\Chrome\Application\chrome.exe',
    r'C:\Program Files (x86)\Google\Chrome\Application\chrome.exe',
    '/Applications/Microsoft Edge.app/Contents/MacOS/Microsoft Edge',
    '/Applications/Google Chrome.app/Contents/MacOS/Google Chrome',
)
_PDF_CHROME_NAMES = ('msedge', 'chrome', 'google-chrome', 'google-chrome-stable',
                     'chromium', 'chromium-browser')


def _find_chrome():
    """Edge/Chrome の実行パスを探す(見つからなければ None)。"""
    import shutil
    for _n in _PDF_CHROME_NAMES:
        _p = shutil.which(_n)
        if _p:
            return _p
    for _p in _PDF_CHROME_CANDIDATES:
        try:
            if Path(_p).exists():
                return _p
        except Exception:
            continue
    return None


def html_to_pdf(html_path, pdf_path, timeout=120, landscape=True):
    """HTMLファイルをPDFに変換。成功したバックエンド名を返す(失敗時 None)。
    ★v136_007: landscape=False で縦向き(読者向け note 版)。Chrome/Edge は
    HTML側の @page 指定に従うので、この引数は Playwright/wkhtmltopdf 用。"""
    import subprocess
    import tempfile
    html_path = Path(str(html_path))
    pdf_path = Path(str(pdf_path))
    _url = html_path.resolve().as_uri()

    # ── 1) Edge / Chrome ヘッドレス ──
    _exe = _find_chrome()
    if _exe:
        # ★--print-to-pdf は user-data-dir が既存プロファイルと衝突すると
        #   無言で失敗することがあるので、必ず一時プロファイルを与える。
        with tempfile.TemporaryDirectory() as _td:
            cmd = [_exe, '--headless=new', '--disable-gpu', '--no-sandbox',
                   '--no-first-run', '--no-pdf-header-footer',
                   f'--user-data-dir={_td}',
                   f'--print-to-pdf={str(pdf_path)}', _url]
            try:
                subprocess.run(cmd, capture_output=True, timeout=timeout)
                if pdf_path.exists() and pdf_path.stat().st_size > 1000:
                    return 'chrome/edge'
                # 旧版Chromeは --headless=new を解さないので従来フラグで再試行
                cmd[1] = '--headless'
                subprocess.run(cmd, capture_output=True, timeout=timeout)
                if pdf_path.exists() and pdf_path.stat().st_size > 1000:
                    return 'chrome/edge'
            except Exception:
                pass

    # ── 2) Playwright ──
    try:
        from playwright.sync_api import sync_playwright
        with sync_playwright() as pw:
            br = pw.chromium.launch()
            pg = br.new_page()
            pg.goto(_url, wait_until='load')
            pg.pdf(path=str(pdf_path), format='A4', landscape=bool(landscape),
                   print_background=True,
                   margin={'top': '8mm', 'bottom': '8mm',
                           'left': '8mm', 'right': '8mm'})
            br.close()
        if pdf_path.exists() and pdf_path.stat().st_size > 1000:
            return 'playwright'
    except Exception:
        pass

    # ── 3) wkhtmltopdf ──
    try:
        import shutil
        _wk = shutil.which('wkhtmltopdf')
        if _wk:
            subprocess.run([_wk, '--quiet', '--orientation',
                            'Landscape' if landscape else 'Portrait',
                            '--page-size', 'A4', '--encoding', 'utf-8',
                            '--margin-top', '8mm', '--margin-bottom', '8mm',
                            '--margin-left', '8mm', '--margin-right', '8mm',
                            str(html_path), str(pdf_path)],
                           capture_output=True, timeout=timeout)
            if pdf_path.exists() and pdf_path.stat().st_size > 1000:
                return 'wkhtmltopdf'
    except Exception:
        pass

    # ── 4) WeasyPrint ──
    try:
        from weasyprint import HTML as _WHTML
        _WHTML(filename=str(html_path)).write_pdf(str(pdf_path))
        if pdf_path.exists() and pdf_path.stat().st_size > 1000:
            return 'weasyprint'
    except Exception:
        pass

    return None


def build_html_nar(html_races, src_name):
    """(name, analysis, extras) のリストから中央競馬スタイルのHTML予想表を組む。
    extras: {ban: dict(waku, seirei, kinryo, kishu_idx)}"""
    # ★【バグ修正 v003】保険としてHTML側でも (場所, レース番号) 昇順に並べ替える。
    #   name = (場所, 距離, レース, 出走時刻)。'1R','10R' のような文字列比較では
    #   1R の次に 10R が来てしまうため、数値を抽出して比較する。
    def _race_sort_key(item):
        name = item[0]
        venue = str(name[0]) if len(name) > 0 else ''
        rlabel = (str(name[2]) if len(name) > 2 else '').translate(_Z2H_DIGITS)
        m = _re.search(r'([0-9]+)', rlabel)
        return (venue, int(m.group(1)) if m else 0, str(name[3]) if len(name) > 3 else '')
    html_races = sorted(html_races, key=_race_sort_key)

    n_gz = sum(1 for _, a, _x in html_races if a.get('shobu'))
    parts = [_NAR_HTML_HEAD.replace('<meta charset="utf-8">',
                                    '<meta charset="utf-8">' + bado_version_meta(), 1)
             .replace('__SRC__', _h(src_name))
             .replace('__NR__', str(len(html_races)))
             .replace('__NGZ__', str(n_gz))
             .replace('__GZ_LEGEND__', _gz_legend_html())]
    for name, analysis, extras in html_races:
        tbl = analysis['table']
        if tbl is None or tbl.empty:
            continue
        venue, dist, rno, ptime = name[0], name[1], name[2], name[3]
        gz_ut = analysis.get('gz_umatan_cond', {})
        gz_um = analysis.get('gz_umaren_cond', {})
        gz_wd = analysis.get('gz_wide_cond', {})   # ★v136_001
        stakes = analysis.get('gz_stakes', {})
        force_ref = analysis.get('force_reference', False)
        active = analysis.get('gz_active_conds', [])
        has_gz = bool(active) and not force_ref
        # 馬表
        rows = []
        for _, row in tbl.iterrows():
            ban = int(row['番']) if pd.notna(row['番']) else 0
            ex = extras.get(ban, {})
            waku = ex.get('waku') or 0
            bg, fg = _NAR_WAKU_STYLE.get(waku, ("#eee", "#000"))
            mk = str(row.get('印', '') or '')
            ev = float(row.get('単勝期待値', 0) or 0)
            wp = float(row.get('単勝確率', 0) or 0)
            fp = float(row.get('複勝確率', 0) or 0)
            himo = row.get('紐馬指数', '')
            try:
                ens_v = float(row.get('単指数', '') or 0)
                ens_txt = f'{ens_v:.1f}'
            except Exception:
                ens_txt = ''
            # ★v137_001: 前売(補正前単勝オッズ)・人気O(人気オッズ)。単勝率がオッズ勝率を上回る馬は前売を緑に。
            def _otxt(v):
                try:
                    v = float(v)
                    return f'{v:.1f}' if v == v and v > 0 else '-'
                except Exception:
                    return '-'
            _mae_txt = _otxt(row.get('前売オッズ'))
            _pop_txt = _otxt(row.get('人気オッズ'))
            try:
                _val_hi = wp >= 1.15 * float(row.get('オッズ勝率')) and wp >= 5.0
            except Exception:
                _val_hi = False
            # 騎手指数60+で騎手名に色マーク
            _ki = ex.get('kishu_idx', '')
            try:
                _ki_hot = (float(_ki) >= 60) if _ki != '' else False
            except Exception:
                _ki_hot = False
            jk_cls = ' jk-hot' if _ki_hot else ''
            _is_ana = ('穴' in mk)                      # ★011 穴条件該当馬
            # 各色マーク
            tan_cls = 'hi-tan' if wp >= 44 else 'prob tan'      # 単勝確率44%+(尖らせ前の40%+相当)
            fuku_cls = 'hi-fuku' if fp >= 74.5 else 'prob fuku'   # 複勝確率74.5%+(尖らせ前の70%+相当)
            idx_cls = ' class="hi-idx"' if _val_hi else ''       # 単勝率 >= オッズ勝率×1.15
            ev_cls = ' class="evhi"' if ev >= 110 else ''        # 色マーク付き=太字(110+のみ)
            _tr_cls = ('axis' if mk == '◎' else '') + (' anaba' if _is_ana else '')
            # ★f文字列の式部分にバックスラッシュを書くと Python 3.11 以前では
            #   SyntaxError(=起動不能)になるため、事前に組み立てておく。
            _ana_cls = ' class="hi-ana"' if _is_ana else ''
            _mk_col = '#e67e22' if _is_ana else _NAR_MARK_COLOR.get(mk, '#000')
            # オッズ欠損(0)は '0.0倍' ではなく '-' と表示する
            try:
                _od_v = float(row["単勝オッズ"])
            except Exception:
                _od_v = 0.0
            _od_txt = f'{_od_v:.1f}' if _od_v > 0 else '-'
            # ★v135_005: 表示専用 補正騎手指数(偏差値)。60+を強調。
            _kdev = ex.get('kishu_dev', '')
            try:
                _kdev_v = float(_kdev)
                _kdev_txt = f'{_kdev_v:.1f}'
                _kdev_cls = 'kdev hi' if _kdev_v >= 60 else 'kdev'
            except Exception:
                _kdev_txt, _kdev_cls = '-', 'kdev'
            rows.append(
                f'<tr class="{_tr_cls.strip()}">'
                f'<td class="mk" style="color:{_mk_col}">{mk}</td>'
                f'<td class="waku" style="background:{bg};color:{fg}">{waku if waku else ""}</td>'
                f'<td>{ban}</td>'
                f'<td class="name">{_h(row.get("馬名", ""))}</td>'
                f'<td>{_h(ex.get("seirei", ""))}</td>'
                f'<td>{_h(ex.get("kinryo", ""))}</td>'
                f'<td class="jk{jk_cls}">{_h(row.get("騎手", ""))}</td>'
                f'<td class="{_kdev_cls}">{_kdev_txt}</td>'
                f'<td>{_h(row.get("展開", ""))}</td>'
                f'<td>{int(row["単勝人気"]) if pd.notna(row.get("単勝人気")) else "-"}</td>'
                f'<td>{_od_txt}</td>'
                f'<td class="{tan_cls}">{wp:.1f}</td>'
                f'<td class="{fuku_cls}">{fp:.1f}</td>'
                f'<td{ev_cls}>{ev:.0f}</td>'
                f'<td{idx_cls}>{_mae_txt}</td>'
                f'<td{_ana_cls}>{_pop_txt}</td>'
                f'<td>{ens_txt}</td>'
                f'<td>{himo}</td>{_stat_td(row)}</tr>')
        # 買い目チップ
        def _chips(recs, cond_map, kind, arrow, tier=None):
            items = []
            for rec in recs:
                k = (rec.ban1, rec.ban2)
                cn = cond_map.get(k, '')
                if tier is not None and gz_tier(cn) != tier:   # ★v135_011 枠フィルタ
                    continue
                lab = _NAR_GZ_LABEL.get(cn)
                st = stakes.get((rec.ban1, rec.ban2, kind), 0)
                if force_ref:
                    tail = '<span class="skipnote">参考(見送り)</span>'
                elif st > 0:
                    tail = f'<span class="stake">{st}円 ★採用</span>'
                else:
                    tail = '<span class="skipnote">—</span>'
                chip = (f'<span class="condchip" style="background:{lab[2]};color:{lab[3]}">'
                        f'{lab[0]}</span>' if lab else '')
                bgc = lab[2] if (lab and not force_ref) else '#eee'
                items.append(
                    f'<span class="bi" style="background:{bgc}"><b>{rec.ban1}{arrow}{rec.ban2}</b>'
                    f'<small>{rec.odds:.1f}倍/的中{rec.hit_rate:.1f}%/EV{rec.ev:.0f}</small>'
                    f'{chip}{tail}{_stat_ura_badge(analysis, cn)}</span>')
            return ''.join(items)
        buys = []
        if has_gz:
            buys.append(f'<div class="gzbanner">◆◆ 採用{len(GZ_COND_DEFS)}条件 買い目'
                        f'（均等買い1点{GZ_FLAT_UNIT}円・本線枠/紐カバー枠は別上限）◆◆</div>')
            for _tk in ('main', 'cover'):
                _acts = [c for c in active if gz_tier(c) == _tk and c in _NAR_GZ_LABEL]
                if not _acts:
                    continue
                labs = ' ／ '.join(f'{_NAR_GZ_LABEL[c][0]} {_NAR_GZ_LABEL[c][1]}' for c in _acts)
                buys.append(f'<div class="gzactive">発火条件【{GZ_TIER_LABEL[_tk]}】: {labs}</div>')
        ut = analysis.get('umatan_recs', [])
        um = analysis.get('umaren_recs', [])
        wd = analysis.get('wide_recs', [])          # ★v136_001
        # ★v135_011: 本線枠 → 紐カバー枠 の順に、行を分けて表示する。
        #   紐カバー枠は「低的中率・高配当依存/下限不足」の参考枠であることを行頭に明記。
        for _tk in ('main', 'cover'):
            _c_ut = _chips(ut, gz_ut, "umatan", "→", _tk)
            _c_um = _chips(um, gz_um, "umaren", "-", _tk)
            _c_wd = _chips(wd, gz_wd, "wide", "-", _tk)     # ★v136_001
            if not (_c_ut or _c_um or _c_wd):
                continue
            _btcls = 'bt' if _tk == 'main' else 'bt anatier'
            # ★v136_001: 表示順は ワイド → 馬連 → 馬単(的中率の高い券種を上に置く)。
            if _c_wd:
                buys.append(f'<div class="brow"><span class="{_btcls}">'
                            f'ワイド【{GZ_TIER_LABEL[_tk]}】</span>{_c_wd}</div>')
            if _c_um:
                buys.append(f'<div class="brow"><span class="{_btcls}">'
                            f'馬連【{GZ_TIER_LABEL[_tk]}】</span>{_c_um}</div>')
            if _c_ut:
                buys.append(f'<div class="brow"><span class="{_btcls}">'
                            f'馬単【{GZ_TIER_LABEL[_tk]}】</span>{_c_ut}</div>')
        # ★v134: 安定運用(トリガミ覚悟)の単勝/複勝
        _st_tan = analysis.get('stable_tan_rec')
        _st_fuku = analysis.get('stable_fuku_rec')
        if _st_tan is not None or _st_fuku is not None:
            def _stable_chip(rec, tag, kind):
                if rec is None:
                    return ''
                _in = tag.startswith('安定')   # 「参考(安定外)」を誤検出しない
                _bg = '#1b7a3a' if _in else '#8a8f98'
                _od = (f'{rec.odds:.1f}倍' if rec.odds and rec.odds > 0 else '—')
                _cls = 'stbet in' if _in else 'stbet out'
                return (f'<span class="{_cls}">'
                        f'<span class="stban">{rec.ban}</span>'
                        f'<span class="stnum">的中<b>{rec.hit_rate:.0f}%</b>'
                        f'<span class="stod">{_od}</span></span>'
                        f'<span class="sttag" style="background:{_bg}">{tag}</span>'
                        f'</span>')
            if _st_tan is not None:
                buys.append('<div class="brow"><span class="bt">単勝【安定】</span>'
                            + _stable_chip(_st_tan, analysis.get('stable_tan_tag', ''), 'tan')
                            + '</div>')
            if _st_fuku is not None:
                buys.append('<div class="brow"><span class="bt">複勝【安定】</span>'
                            + _stable_chip(_st_fuku, analysis.get('stable_fuku_tag', ''), 'fuku')
                            + '</div>')
        # ★v135_028: 旧・次点本線枠A/B/C は停止(SUBLINE_LEGACY_ENABLE=False)。
        #   subline_rec / subline_b_rec / subline_c_rec は常に None のため描画されない。
        # ★v135_028: 次点本線枠(勝ち筋マップ)= 本線(MB1/MB2)非採用の sign通過条件。
        #   表示専用(実弾ではない)。同じ買い目は本線へ集約済みで、ここには本線と
        #   重複しない買い目のみが残る。CI下限100%未満のため検証候補。
        _KIND_JP = {'umatan': '馬単', 'umaren': '馬連', 'wide': 'ワイド'}
        _KIND_SEP = {'umatan': '→', 'umaren': '-', 'wide': '-'}
        for _jr in (analysis.get('jiten_recs') or []):
            _jkn = _KIND_JP.get(_jr['kind'], _jr['kind'])
            _jsep = _KIND_SEP.get(_jr['kind'], '-')
            _j_buy = ' / '.join('%d%s%d' % (_jr['axis_ban'], _jsep, _p['ban'])
                                for _p in _jr['partners'])
            _j_det = '・'.join(
                '%d番(複勝%.0f%%%s)'
                % (_p['ban'], _p['place'],
                   '/%.1f倍' % _p['odds'] if _p.get('odds') else '')
                for _p in _jr['partners'])
            # ★v135_035: おすすめ条件ランキング由来は順位・タイプ・回収率を表示。
            #   (それ以外は従来どおり 検ROI/的中/N と CI下限の注意書きを表示)
            if _jr.get('src') == 'v137':      # ★v137_001: 新条件(検証前)
                _j_stat = '%s〔%s・実弾外〕 未検証' % (_jr['tag'], _jr.get('typ', ''))
            elif _jr.get('src') == 'sheet':     # ★v136_018: 実戦条件シートの観察条件
                _j_stat = ('%s〔%s・実弾外〕 検ROI%.1f/的中%.1f%%/検証%dR(7/16〜9/23)'
                           % (_jr['tag'], _jr.get('typ', ''),
                              _jr['roi'], _jr['hit'], _jr['n']))
            elif _jr.get('src') == 'osusume':
                _j_stat = ('%s〔%s〕 回収%.1f%%/的中%.1f%%/N%d'
                           % (_jr['tag'], _jr.get('typ', ''),
                              _jr['roi'], _jr['hit'], _jr['n']))
            else:
                _j_stat = ('%s 検ROI%.1f/的中%.1f%%/N%d・CI下限100%%未満'
                           % (_jr['tag'], _jr['roi'], _jr['hit'], _jr['n']))
            _j_note = ('<span class="subnote">%s × %s / %s'
                       '<span class="subwarn">%s</span>'
                       '</span>'
                       % (_jr['axis_label'], _jr['atoms_label'], _j_det, _j_stat))
            buys.append(
                '<div class="brow"><span class="bt subtier">%s【次点本線枠】</span>'
                '<span class="bi"><b>%s</b>%s</span>%s</div>'
                % (_jkn, _j_buy, _stat_ura_badge(analysis, _jr.get('tag')), _j_note))
        # ★v136_003: 参考予想(◎から予想印▲△☆の馬への馬連・ワイド)。実弾ではない。
        _refs = analysis.get('reference_bets') or []
        _REF_KIND_JP = {'umatan': '馬単', 'umaren': '馬連', 'wide': 'ワイド'}
        _REF_KIND_SEP = {'umatan': '→', 'umaren': '-', 'wide': '-'}
        for _rf in _refs:
            _arrow = _REF_KIND_SEP.get(_rf['kind'], '-')
            _kn = _REF_KIND_JP.get(_rf['kind'], _rf['kind'])
            _buy = ' / '.join('%d%s%d' % (_rf['axis_ban'], _arrow, _b)
                              for _b in _rf['partner_bans'])
            _dsc = _h(str(_rf.get('desc') or ''))
            if _rf.get('is_fallback'):
                _note = ('<span style="color:#b45309">%s ※該当馬なしのフォールバック</span>'
                         % _dsc)
            else:
                _note = ('<span style="color:#6b7280">%s</span>' % _dsc)
            if _rf.get('shobu'):
                buys.append('<div class="brow">'
                            '<span class="bt" style="background:#c00000">'
                            '%s【勝負】</span>'
                            '<b>%s</b> %s &nbsp;%s</div>'
                            % (_kn, _h(_rf['name']), _buy, _note))
            else:
                buys.append('<div class="brow" style="opacity:.85">'
                            '<span class="bt" style="background:#9aa0a6">'
                            '%s【参考】</span>'
                            '<b>%s</b> %s &nbsp;%s</div>'
                            % (_kn, _h(_rf['name']), _buy, _note))
        if (not ut and not um and not wd and _st_tan is None and _st_fuku is None
                and not _refs):
            buys.append('<div class="skip">― 買い目なし ―</div>')
        rel = analysis.get('judgment_class', '')
        # 軸判定を大きく強調(S/A/B/C/D級軸を色付きバッジで)
        _axc = str(analysis.get('axis_class', ''))
        _gl = 'D'
        for _g in ('S', 'A', 'B', 'C', 'D'):
            if _axc.startswith(_g + '級'):
                _gl = _g; break
        grade_badge = f'<span class="gradebig g-{_gl}">{_axc}</span>'
        conf = ''
        try:
            _rec = calculate_confidence_score(analysis)
            conf = (f'<span class="rel" style="background:#{_rec["bg_hex"]};color:#{_rec["font_hex"]}">'
                    f'信頼度{_rec["score"]} {_rec["category"]}</span>')
        except Exception:
            pass
        _shobu = bool(analysis.get('shobu'))
        badge = ('<span class="badge">★%s(単指数%.1f)</span>' % (SHOBU_LABEL, float(analysis.get('shobu_idx') or 0))) if _shobu else ''
        # race_cat(レース分類)部分のみ rmeta に(級はバッジで表示済み)
        _race_cat = rel.split(' / ')[-1] if ' / ' in rel else rel
        parts.append(f"""
        <div class="{'race gz-race' if _shobu else 'race'}">
          <div class="rhead"><span class="rtitle">{_h(venue)} {_h(rno)}R</span>{grade_badge}{conf}
            <span class="rmeta">{_h(_dist_txt(dist))} / 発走{_h(ptime or '-')} / {len(tbl)}頭 / {_h(_race_cat)}</span>{badge}</div>
          <table class="grid"><thead><tr>
            <th>印</th><th>枠</th><th>馬番</th><th>馬名</th><th>性齢</th><th>斤量</th>
            <th>騎手</th><th>{KISHU_DEV_LABEL}</th><th>脚質</th><th>人気</th><th>単オッズ</th>
            <th>単勝率</th><th>複勝率</th><th>期待値</th><th>前売</th><th>人気O</th><th>単指数</th><th>紐馬指数</th><th>統計印</th><th>統計%</th><th>オッズ%</th>
          </tr></thead><tbody>{''.join(rows)}</tbody></table>
          <div class="buys">{''.join(buys)}</div>{_data_note_html(analysis)}{_stat_block_html(analysis)}</div>""")
    parts.append('<div class="footer">v137_001 オッズ×統計 アンサンブル版。勝負レース=◎の単指数90以上。買い目は表示のみ(実弾なし)。損失許容の範囲で。</div></body></html>')
    return '\n'.join(parts)


# ══════════════════════════════════════════════════════════════════
# ★v136_007: note読者向け 予想表(HTML/PDF)
#   ・既存の予想表(build_html_nar)は研究用データ(keiba_tool の読込対象)として
#     そのまま残し、読者向けは別ファイル(*_note.html / *_note.pdf)で出力する。
#     既存HTMLの構造を変えると keiba_tool の読み込みが壊れるため。
#   ・内部用の情報(賭け金・点数上限・発火条件の説明・次点枠など)は出さない。
#   ・縦長A4/スマホ幅で読める1カラム構成。冒頭に「本日の推奨買い目」を発走順で一覧化。
# ══════════════════════════════════════════════════════════════════
NOTE_HTML_ENABLE = True          # False にすると読者向けHTML/PDFを出力しない
NOTE_SHOW_REFERENCE = True       # 参考買い目(◎→印の馬)を表示する
NOTE_SHOW_STABLE_OUT = False     # 単勝/複勝の「参考(安定外)」も表示する(既定=安定のみ)
NOTE_SHOW_COND_CODE = False      # True で推奨買い目に条件コード(V1等)を小さく添える(成績を条件別に公開する時用)
NOTE_BRAND = 'BADO'
NOTE_DISCLAIMER = ('掲載している確率・指数・買い目は、単勝オッズと過去成績の統計から計算した予測です。'
                   '的中や回収を保証するものではありません。馬券の購入はご自身の判断で、'
                   '無理のない範囲でお楽しみください。')

_NOTE_WEEK = '月火水木金土日'
_NOTE_MARK_STYLE = {   # (文字色, 背景色)
    '◎': ('#ffffff', '#c62839'), '○': ('#ffffff', '#1f5fbf'), '▲': ('#ffffff', '#0b8a64'),
    '△': ('#ffffff', '#b8791a'), '☆': ('#ffffff', '#7a4bd3'), '穴': ('#ffffff', '#d4621a'),
    '注': ('#ffffff', '#7d6608'),
}
_NOTE_GRADE_STYLE = {   # 軸の信頼度バッジ (文字色, 背景色)
    'S': ('#3b2a00', '#f4c542'), 'A': ('#ffffff', '#e0673a'), 'B': ('#ffffff', '#3b82c4'),
    'C': ('#ffffff', '#3a9a5b'), 'D': ('#ffffff', '#8a929b'),
}

_NOTE_CSS = """
:root{--ink:#1c2430;--ink2:#4b5563;--ink3:#7b8491;--line:#dfe3e8;--paper:#ffffff;--bg:#eef1f4;
 --brand:#0c3b2e;--brand2:#135c47;--gold:#c9a24a;--win:#d0364a;--plc:#2563c4;--good:#127a53;
 --soft:#f6f8fa;--pick:#fff7e6;--pickline:#e9b949}
*{box-sizing:border-box}
html,body{margin:0;background:var(--bg);color:var(--ink);
 font-family:"BIZ UDPGothic","Yu Gothic UI","Yu Gothic","Hiragino Sans","Meiryo",sans-serif;
 -webkit-print-color-adjust:exact;print-color-adjust:exact}
.wrap{max-width:860px;margin:0 auto;padding:0 12px 28px}
.num{font-variant-numeric:tabular-nums}
/* 表紙 */
.cover{background:var(--brand);color:#fff;padding:22px 20px 18px;border-radius:0 0 14px 14px;margin:0 0 16px}
.cover .brand{font-size:12px;letter-spacing:.2em;color:var(--gold);font-weight:700}
.cover h1{margin:6px 0 2px;font-size:26px;letter-spacing:.02em}
.cover .venues{font-size:14px;opacity:.9}
.kpis{display:grid;grid-template-columns:repeat(3,1fr);gap:8px;margin-top:14px}
.kpi{background:rgba(255,255,255,.1);border-radius:10px;padding:8px 10px}
.kpi b{display:block;font-size:22px;font-variant-numeric:tabular-nums}
.kpi span{font-size:11px;opacity:.85}
/* セクション見出し */
h2{font-size:17px;margin:22px 2px 10px;display:flex;align-items:center;gap:8px}
h2::before{content:"";width:5px;height:18px;background:var(--brand2);border-radius:3px}
/* 本日の推奨買い目 */
.index{background:var(--paper);border:1px solid var(--line);border-radius:12px;overflow:hidden}
.index table{width:100%;border-collapse:collapse;font-size:13.5px}
.index th{background:var(--soft);color:var(--ink2);font-size:11.5px;font-weight:700;text-align:left;padding:7px 10px;border-bottom:1px solid var(--line)}
.index td{padding:8px 10px;border-bottom:1px solid var(--line);vertical-align:top}
.index tr:last-child td{border-bottom:0}
.index .t{font-weight:700;white-space:nowrap}
.index .race{white-space:nowrap}
.index .race a{color:var(--ink);text-decoration:none;font-weight:700}
.index .chips{display:flex;flex-wrap:wrap;gap:5px}
.chip{display:inline-flex;align-items:center;gap:5px;border:1px solid var(--pickline);background:var(--pick);border-radius:7px;padding:2px 8px 2px 3px;white-space:nowrap}
.chip .ty{font-size:10.5px;font-weight:700;color:#fff;background:#b7791f;border-radius:4px;padding:1px 5px}
.chip b{font-size:14.5px;font-variant-numeric:tabular-nums}
.index .ax{color:var(--ink2);font-size:12px;white-space:nowrap}
.empty{background:var(--paper);border:1px dashed var(--line);border-radius:12px;padding:16px;text-align:center;color:var(--ink3)}
/* 見方 */
.guide{background:var(--paper);border:1px solid var(--line);border-radius:12px;padding:12px 14px;font-size:12.5px;color:var(--ink2);line-height:1.7}
.guide dl{display:grid;grid-template-columns:auto 1fr;gap:4px 12px;margin:0}
.guide dt{font-weight:700;color:var(--ink);white-space:nowrap}
.guide dd{margin:0}
.marks{display:flex;flex-wrap:wrap;gap:6px 12px}
/* レースカード */
.card{background:var(--paper);border:1px solid var(--line);border-radius:12px;margin:12px 0;overflow:hidden;break-inside:avoid;page-break-inside:avoid}
.card.pick{border:2px solid var(--pickline)}
.chead{display:flex;align-items:center;flex-wrap:wrap;gap:4px 10px;padding:10px 12px;border-bottom:1px solid var(--line)}
.chead .ttl{display:flex;align-items:baseline;flex-wrap:wrap;gap:2px 10px}
.chead .rn{font-size:19px;font-weight:800;white-space:nowrap}
.chead .meta{font-size:12.5px;color:var(--ink2)}
.chead .right{margin-left:auto;display:flex;align-items:center;gap:6px}
.grade{display:inline-flex;align-items:center;gap:6px;font-size:12px;color:var(--ink2);white-space:nowrap}
.grade i{font-style:normal;display:inline-flex;align-items:center;justify-content:center;width:26px;height:26px;border-radius:7px;font-weight:800;font-size:15px}
.flag{font-size:11.5px;font-weight:700;color:#7a4d00;background:#fde7b0;border-radius:999px;padding:3px 9px;white-space:nowrap}
.cmt{font-size:12.5px;color:var(--ink2);padding:7px 12px;background:var(--soft);border-bottom:1px solid var(--line)}
table.h{width:100%;border-collapse:collapse;font-size:13px}
.h th{font-size:10.5px;color:var(--ink3);font-weight:700;padding:5px 4px;border-bottom:1px solid var(--line);text-align:center;white-space:nowrap}
.h td{padding:6px 4px;border-bottom:1px solid #eef0f3;text-align:center;vertical-align:middle}
.h tr:last-child td{border-bottom:0}
.h tr.hl td{background:#fffaf0}
.h td.l{text-align:left}
.mk{display:inline-flex;align-items:center;justify-content:center;width:22px;height:22px;border-radius:50%;font-weight:800;font-size:13px}
.ban{display:inline-flex;align-items:center;justify-content:center;min-width:24px;height:24px;border-radius:5px;font-weight:800;font-size:13px;border:1px solid rgba(0,0,0,.18)}
.hn{font-weight:700;font-size:14px}
.hs{font-size:11px;color:var(--ink3);margin-top:1px}
.bar{position:relative;display:block;width:62px;height:16px;background:#eef1f5;border-radius:4px;overflow:hidden;margin:0 auto}
.bar>s{position:absolute;left:0;top:0;bottom:0;border-radius:4px}
.bar>em{position:relative;font-style:normal;font-size:11.5px;font-weight:700;line-height:16px;font-variant-numeric:tabular-nums}
.bar.w>s{background:rgba(208,54,74,.28)}.bar.p>s{background:rgba(37,99,196,.24)}
.hscroll{overflow-x:auto;-webkit-overflow-scrolling:touch}
.h thead th{background:var(--soft)}
.h tr.grp th{font-size:10.5px;color:var(--ink2);border-bottom:1px solid var(--line)}
.h th small{font-weight:400;font-size:9.5px}
.h .g1{border-left:1px solid var(--line)}
.h th.l{text-align:left}
.mb{display:inline-flex;align-items:center;gap:4px}
.mk.none{visibility:hidden}
.mkw{position:relative;display:inline-flex}
.anab{position:absolute;right:-8px;top:-7px;font-size:8.5px;font-weight:800;line-height:1;color:#fff;background:#d4621a;border:1px solid #fff;border-radius:4px;padding:1px 2px}
.sty{font-weight:700;font-size:12.5px}
.s-逃{color:#c0392b}.s-先{color:#b8660b}.s-差{color:#1f6fd6}.s-追{color:#6b46c1}
.jkhot{color:#c0392b;font-weight:700}
.hot{color:var(--good);font-weight:800}
.anav{color:#d4621a;font-weight:800}
.h tr.ana td{background:#fff6ee}
.bar.w.hi>s{background:rgba(208,54,74,.55)}.bar.p.hi>s{background:rgba(37,99,196,.5)}
.scrollhint{display:none}
/* 買い目 */
.buys{padding:10px 12px 12px;border-top:1px solid var(--line);display:grid;gap:8px}
.bgrp .bl{font-size:11.5px;font-weight:700;color:var(--ink2);margin-bottom:5px}
.bgrp.main .bl{color:#8a5a00}
.tks{display:flex;flex-wrap:wrap;gap:6px}
.tk{display:inline-flex;align-items:center;gap:7px;border-radius:9px;padding:5px 10px;border:1px solid var(--line);background:var(--soft)}
.bgrp.main .tk{background:var(--pick);border-color:var(--pickline)}
.tk .ty{font-size:11px;font-weight:700;color:#fff;background:var(--brand2);border-radius:5px;padding:1px 6px}
.bgrp.main .tk .ty{background:#b7791f}
.tk .cb{font-size:17px;font-weight:800;font-variant-numeric:tabular-nums;letter-spacing:.02em}
.tk .nm{font-size:11px;color:var(--ink2)}
.tk .cd{font-size:10px;color:var(--ink3)}
.bgrp.ref .tk .cb{font-size:15px}
.none{font-size:12.5px;color:var(--ink3)}
.foot .ver{margin-top:6px;font-size:10px;color:#a0a7b1}
.foot{margin-top:22px;font-size:11.5px;color:var(--ink3);line-height:1.7;border-top:1px solid var(--line);padding-top:12px}
@media (max-width:560px){
 .col-ext{display:none}
 .h .sk{position:sticky;background:var(--paper);z-index:1}
 .h thead .sk{background:var(--soft);z-index:2}
 .h tr.hl td.sk{background:#fffaf0}.h tr.ana td.sk{background:#fff6ee}
 .h .sk1{left:0;min-width:52px}
 .h .sk2{left:52px;min-width:112px;max-width:132px;box-shadow:4px 0 5px -4px rgba(0,0,0,.25)}
 .hs{white-space:normal}
 .scrollhint{display:block;font-size:10.5px;color:var(--ink3);text-align:right;padding:3px 10px 0}
 .wrap{padding:0 8px 24px}
 .cover{margin:0 -8px 14px;padding:18px 16px 14px}
 .cover h1{font-size:22px}
 .kpi{padding:7px 8px}.kpi b{font-size:20px}.kpi span{font-size:10.5px}
 .bar{width:42px}
 .bar>em{font-size:11px}
 .h{font-size:12.5px}
 .h th{font-size:10px;padding:5px 2px}
 .h td{padding:6px 2px}
 .mk{width:20px;height:20px;font-size:12px}
 .ban{min-width:22px;height:22px}
 .hn{font-size:13.5px}
 .chead .right{margin-left:0}
 .index td{padding:7px 7px}
 .index .ax{display:none}
}
@page{size:A4 portrait;margin:9mm}
@media print{
 html,body{background:#fff}
 .wrap{max-width:none;padding:0}
 .cover{margin:0 0 12px;border-radius:10px}
 .card,.guide{break-inside:avoid;page-break-inside:avoid}
 .index{overflow:visible}
 .index tr{break-inside:avoid;page-break-inside:avoid}
 .index thead{display:table-header-group}
 h2{break-after:avoid;page-break-after:avoid}
 /* 1ページに2レース入るよう詰める */
 .card{margin:8px 0}
 .chead{padding:6px 10px}.chead .rn{font-size:17px}
 .cmt{padding:4px 10px;font-size:11.5px}
 .h{font-size:11px}
 .h td{padding:3px 3px}.h th{padding:3px 3px;font-size:9.5px}
 .bar{width:44px}
 .hscroll{overflow:visible}
 .hn{font-size:12.5px;display:inline}.hs{display:inline;margin-left:6px}
 .mk{width:18px;height:18px;font-size:11px}.ban{min-width:20px;height:20px;font-size:12px}
 .bar{height:14px}.bar>em{line-height:14px;font-size:10.5px}
 .buys{padding:6px 10px 8px;gap:5px}
 .tk{padding:3px 8px}.tk .cb{font-size:15px}
}
"""


def _note_date_label(src_name):
    """予想対象日を 'YYYY年M月D日(曜)' で返す。取れなければ空文字。"""
    _d = None
    try:
        if CURRENT_RACECARD_DATE:
            _d = int(CURRENT_RACECARD_DATE)
    except Exception:
        _d = None
    if not _d:
        _m = _re.findall(r'(20\d{6})', str(src_name))
        if _m:
            _d = int(_m[-1])
    if not _d:
        return ''
    try:
        _dt = date(_d // 10000, (_d // 100) % 100, _d % 100)
        return '%d年%d月%d日(%s)' % (_dt.year, _dt.month, _dt.day, _NOTE_WEEK[_dt.weekday()])
    except Exception:
        return ''


def _note_rno(rno):
    _s = str(rno).translate(_Z2H_DIGITS)
    _m = _re.search(r'([0-9]+)', _s)
    return int(_m.group(1)) if _m else 0


def _note_mark(mk):
    mk = str(mk or '').strip()
    if not mk:
        return ''
    # ★v136_014: 『○穴』『▲穴』等は ○▲△ を残して小さな『穴』バッジを併記する
    #   (以前は『穴』だけになり、参考買い目の相手 ○▲△ が読み取れなかった)。
    _base = mk.replace('穴', '').strip()
    if not _base:
        fg, bg = _NOTE_MARK_STYLE['穴']
        return '<span class="mk" style="color:%s;background:%s">穴</span>' % (fg, bg)
    _key = _base[:1]
    fg, bg = _NOTE_MARK_STYLE.get(_key, ('#ffffff', '#6b7280'))
    _html = '<span class="mk" style="color:%s;background:%s">%s</span>' % (fg, bg, _h(_key))
    if '穴' in mk:
        _html = '<span class="mkw">%s<span class="anab">穴</span></span>' % _html
    return _html


def _note_bar(val, cls, scale=100.0):
    try:
        v = float(val)
    except Exception:
        return ''
    w = max(0.0, min(100.0, v / scale * 100.0))
    return ('<span class="bar %s"><s style="width:%.0f%%"></s><em>%.1f</em></span>'
            % (cls, w, v))


def build_html_note(html_races, src_name):
    """note読者向けの予想表HTMLを組む(★v136_007)。(name, analysis, extras) のリストを受け取る。
    買い目・印・確率などの中身は build_html_nar と同じ analysis から取り、表示だけを変える。"""
    def _race_sort_key(item):
        name = item[0]
        return (str(name[0]), _note_rno(name[2] if len(name) > 2 else ''),
                str(name[3]) if len(name) > 3 else '')
    races = sorted(html_races, key=_race_sort_key)

    date_label = _note_date_label(src_name)
    venues = []
    for name, _a, _x in races:
        if name[0] not in venues:
            venues.append(name[0])

    index_rows = []     # (発走, 場, R, anchor, [(券種, 組番)], ◎馬名)
    cards = []
    n_pick_races = 0
    n_pick_pts = 0

    for name, analysis, extras in races:
        tbl = analysis.get('table')
        if tbl is None or tbl.empty:
            continue
        venue, dist, rno, ptime = name[0], name[1], name[2], name[3]
        rnum = _note_rno(rno)
        anchor = 'r-%s-%d' % (_re.sub(r'\W', '', str(venue)) or 'v', rnum)
        ptime_s = str(ptime)[:5] if ptime is not None else ''
        names_by_ban = {}
        for _, row in tbl.iterrows():
            try:
                names_by_ban[int(row['番'])] = str(row.get('馬名', '') or '')
            except Exception:
                pass

        # ── 馬表 ── ★研究用HTMLと同じ全項目を表示(並びは 馬 / オッズ / AI予測 / 指数 の4グループ)
        def _f(v, d=None):
            try:
                x = float(v)
                return d if pd.isna(x) else x
            except Exception:
                return d
        trs = []
        for _, row in tbl.iterrows():
            try:
                ban = int(row['番'])
            except Exception:
                continue
            ex = extras.get(ban, {})
            waku = ex.get('waku') or 0
            wbg, wfg = _NAR_WAKU_STYLE.get(waku, ('#f1f3f5', '#1c2430'))
            mk = str(row.get('印', '') or '')
            is_ana = '穴' in mk
            wp = _f(row.get('単勝確率'), 0.0)
            fp = _f(row.get('複勝確率'), 0.0)
            ev = _f(row.get('単勝期待値'), 0.0)
            # ★v137_001: note版は前売(補正前単勝オッズ)・人気オッズを表示しない(単勝オッズと軸馬/紐馬指数のみ)。
            od = _f(row.get('単勝オッズ'), 0.0)
            ens = _f(row.get('単指数'))
            himo = _f(row.get('紐馬指数'))
            kdev = _f(ex.get('kishu_dev', ''))
            kidx = _f(ex.get('kishu_idx', ''))
            pop = _f(row.get('単勝人気'))
            pop_s = str(int(pop)) if (pop is not None and pop > 0) else '-'
            style = str(row.get('展開', '') or '').strip() or '-'
            jk = str(row.get('騎手', '') or '')
            jk_html = ('<span class="jkhot">%s</span>' % _h(jk)) if (kidx is not None and kidx >= 60) else _h(jk)
            sub = ' '.join(x for x in [_h(str(ex.get('seirei', '') or '')),
                                       _h(str(ex.get('kinryo'))) if ex.get('kinryo') not in (None, '') else '',
                                       jk_html] if x)
            _cls = []
            if mk.strip().startswith(('◎', '○')):
                _cls.append('hl')
            if is_ana:
                _cls.append('ana')
            trs.append(
                '<tr%s>' % ((' class="%s"' % ' '.join(_cls)) if _cls else '')
                + '<td class="sk sk1"><span class="mb">%s<span class="ban" style="background:%s;color:%s">%d</span></span></td>'
                  % (_note_mark(mk) or '<span class="mk none"></span>', wbg, wfg, ban)
                + '<td class="sk sk2 l"><div class="hn">%s</div><div class="hs">%s</div></td>'
                  % (_h(row.get('馬名', '')), sub)
                + '<td class="sty s-%s">%s</td>' % (_h(style[:1]), _h(style))
                + '<td class="num g1">%s</td>' % pop_s
                + '<td class="num">%s</td>' % (('%.1f' % od) if od > 0 else '-')
                + '<td class="g1">%s</td>' % _note_bar(wp, 'w' + (' hi' if wp >= 44 else ''), 66.0)
                + '<td>%s</td>' % _note_bar(fp, 'p' + (' hi' if fp >= 74.5 else ''), 100.0)
                + '<td class="num%s">%.0f</td>' % (' hot' if ev >= 110 else '', ev)
                + '<td class="num g1%s">%s</td>' % (' anav' if is_ana else '',
                                                    ('%.1f' % ens) if ens is not None else '-')
                + '<td class="num">%s</td>' % (('%.0f' % himo) if himo is not None else '-')
                + '<td class="num%s">%s</td>' % (' hot' if (kdev or 0) >= 60 else '',
                                                 ('%.1f' % kdev) if kdev is not None else '-')
                + '</tr>')

        # ── 勝負レース(◎の単指数90以上): ◎→○▲△(馬連・ワイド・馬単各3点。紐は単勝確率順) ──
        main_tks = []
        _shobu = bool(analysis.get('shobu'))
        if _shobu:
            for rf in analysis.get('reference_bets') or []:
                _sep = GZ_KIND_SEP.get(rf.get('kind'), '-')
                for _b in rf.get('partner_bans', []):
                    main_tks.append((rf.get('kind'), rf['axis_ban'], _b, _sep, ''))
        has_main = bool(main_tks)
        if has_main:
            n_pick_races += 1
            n_pick_pts += len(main_tks)
            _ax_nm = ''
            for _, _r in tbl.iterrows():
                if str(_r.get('印', '') or '').strip().startswith('◎'):
                    _ax_nm = '%s %s' % (_r['番'], _r.get('馬名', ''))
                    break
            index_rows.append((ptime_s, venue, rnum, anchor,
                               [(GZ_KIND_JP.get(k, k), '%d%s%d' % (b1, ar, b2))
                                for k, b1, b2, ar, _cn in main_tks], _ax_nm))

        def _tk(kind_jp, combo, nm='', code=''):
            return ('<span class="tk"><span class="ty">%s</span><span class="cb">%s</span>%s%s</span>'
                    % (_h(kind_jp), _h(combo),
                       ('<span class="nm">%s</span>' % _h(nm)) if nm else '',
                       ('<span class="cd">%s</span>' % _h(code)) if code else ''))

        buys = []
        if main_tks:
            _lab = '勝負レース（◎から単勝確率の高い3頭へ）'
            _items = ''.join(
                _tk(GZ_KIND_JP.get(k, k), '%d%s%d' % (b1, ar, b2),
                    '%s・%s' % (names_by_ban.get(b1, ''), names_by_ban.get(b2, '')),
                    (cn if NOTE_SHOW_COND_CODE else ''))
                for k, b1, b2, ar, cn in main_tks)
            buys.append('<div class="bgrp main"><div class="bl">%s</div><div class="tks">%s</div></div>'
                        % (_lab, _items))

        # ── 単勝・複勝(安定) ──
        st_items = []
        for rec_key, tag_key, kind_jp, prob_lab in (('stable_tan_rec', 'stable_tan_tag', '単勝', '勝率'),
                                                    ('stable_fuku_rec', 'stable_fuku_tag', '複勝', '複勝率')):
            rec = analysis.get(rec_key)
            if rec is None:
                continue
            tag = str(analysis.get(tag_key, '') or '')
            if not tag.startswith('安定') and not NOTE_SHOW_STABLE_OUT:
                continue
            nm = '%s ／ %s%.0f%%' % (names_by_ban.get(rec.ban, ''), prob_lab, rec.hit_rate)
            st_items.append(_tk(kind_jp, str(rec.ban), nm))
        if st_items:
            buys.append('<div class="bgrp"><div class="bl">単勝・複勝（的中率重視）</div>'
                        '<div class="tks">%s</div></div>' % ''.join(st_items))

        # ── 参考買い目(◎→印の馬) ──
        if NOTE_SHOW_REFERENCE and not _shobu:
            ref_items = []
            for rf in analysis.get('reference_bets') or []:
                kn = GZ_KIND_JP.get(rf.get('kind'), rf.get('kind'))
                sep = GZ_KIND_SEP.get(rf.get('kind'), '-')
                combo = ' ／ '.join('%d%s%d' % (rf['axis_ban'], sep, b) for b in rf.get('partner_bans', []))
                if combo:
                    ref_items.append(_tk(kn, combo, '◎から印の馬へ'))
            if ref_items:
                buys.append('<div class="bgrp ref"><div class="bl">参考買い目（期待値の高い3頭）</div>'
                            '<div class="tks">%s</div></div>' % ''.join(ref_items))
        if not buys:
            buys.append('<div class="none">このレースは買い目なし</div>')

        # ── レース見出し ──
        _axc = str(analysis.get('axis_class', '') or '')
        _gl = ''
        for _g in ('S', 'A', 'B', 'C', 'D'):
            if _axc.startswith(_g + '級'):
                _gl = _g
                break
        grade_html = ''
        conf_txt = ''
        try:
            _rec = calculate_confidence_score(analysis)
            conf_txt = '信頼度 %s%s' % (_rec['score'], (' ' + str(_rec.get('category', ''))) if _rec.get('category') else '')
        except Exception:
            pass
        if _gl:
            fg, bg = _NOTE_GRADE_STYLE[_gl]
            grade_html = ('<span class="grade"><i style="color:%s;background:%s">%s</i>%s</span>'
                          % (fg, bg, _gl, _h(('%s軸 ・ ' % _gl) + conf_txt)))
        flag = ('<span class="flag">%s</span>' % SHOBU_LABEL) if has_main else ''
        cmt = str(analysis.get('judgment_comment', '') or '').strip()
        _rel = str(analysis.get('judgment_class', '') or '')
        race_cat = (_rel.split(' / ')[-1] if ' / ' in _rel else _rel).strip()
        try:
            dist_s = '%dm' % int(float(dist))
        except Exception:
            dist_s = str(dist)
        cards.append(
            '<section class="card%s" id="%s">' % (' pick' if has_main else '', anchor)
            + '<div class="chead"><span class="ttl"><span class="rn">%s %dR</span>' % (_h(venue), rnum)
            + '<span class="meta">%s発走 ・ %s ・ %d頭%s</span></span>' % (
                _h(ptime_s), _h(dist_s), len(tbl), (' ・ ' + _h(race_cat)) if race_cat else '')
            + '<span class="right">%s%s</span></div>' % (flag, grade_html)
            + ('<div class="cmt">%s</div>' % _h(cmt) if cmt and cmt != 'データなし' else '')
            + '<div class="scrollhint">横にスクロールで全項目 →</div>'
            + '<div class="hscroll"><table class="h"><thead>'
              '<tr class="grp"><th class="sk sk1" rowspan="2">印<br>馬番</th>'
              '<th class="sk sk2 l" rowspan="2">馬名<br><small>性齢 斤量 騎手</small></th>'
              '<th rowspan="2">脚質</th><th colspan="2" class="g1">オッズ</th>'
              '<th colspan="3" class="g1">予測(オッズ×統計)</th><th colspan="3" class="g1">指数</th></tr>'
              '<tr><th class="g1">人気</th><th>単勝</th><th class="g1">勝率%</th><th>複勝率%</th><th>期待値</th>'
              '<th class="g1">単指数</th><th>紐馬</th><th>騎手<br>指数</th></tr></thead><tbody>'
            + ''.join(trs) + '</tbody></table></div>'
            + '<div class="buys">%s</div></section>' % ''.join(buys))

    # ── 本日の勝負レース(発走順) ──
    index_rows.sort(key=lambda t: (t[0] or '99:99', str(t[1]), t[2]))
    if index_rows:
        _ir = ''.join(
            '<tr><td class="t num">%s</td><td class="race"><a href="#%s">%s %dR</a></td>'
            '<td><span class="chips">%s</span></td><td class="ax">◎ %s</td></tr>'
            % (_h(t), a, _h(v), r,
               ''.join('<span class="chip"><span class="ty">%s</span><b>%s</b></span>'
                       % (_h(k), _h(c)) for k, c in tks), _h(n))
            for t, v, r, a, tks, n in index_rows)
        index_html = ('<div class="index"><table><thead><tr><th>発走</th><th>レース</th><th>買い目</th>'
                      '<th class="ax">本命◎</th></tr></thead><tbody>%s</tbody></table></div>' % _ir)
    else:
        index_html = '<div class="empty">本日の勝負レース(単指数90以上)はありません</div>'

    guide_html = (
        '<div class="guide"><dl>'
        '<dt>印</dt><dd><span class="marks">'
        + ' '.join('%s %s' % (_note_mark(k), v) for k, v in
                   ((('◎', '本命'), ('○', '対抗'), ('▲', '単穴'), ('△', '連下'), ('穴', '穴馬')) if REF_NEW_LOGIC
                    else (('◎', '本命'), ('○', '対抗'), ('▲', '単穴'), ('△', '連下'), ('☆', '押さえ'), ('穴', '穴馬'))))
        + '</span></dd>'
        '<dt>脚質</dt><dd>逃=逃げ・先=先行・差=差し・追=追込。</dd>'
        '<dt>騎手の色</dt><dd>騎手名が<span class="jkhot">色付き</span>は、騎手指数が高い(60以上)騎手です。</dd>'
        '<dt>勝率・複勝率</dt><dd>単勝オッズから見た勝率と、出馬表の成績から計算した統計の勝率を合わせて推定した'
        '「1着になる確率」「3着以内に入る確率」(%)。勝率44%以上・複勝率74.5%以上は濃い色で表示。</dd>'
        '<dt>単勝</dt><dd>予想オッズ。前売の単勝オッズと、人気になりそうな要素(騎手・成績・当地実績・最高タイム)から'
        '推定したオッズを合わせたもの。人気はこの順番です。</dd>'
        '<dt>期待値</dt><dd>勝率×予想オッズ。100を超えるほど、オッズに対して割安な馬です。</dd>'
        '<dt>単指数</dt><dd>単勝確率を100点満点に換算した点数。90以上は鉄板クラス(単勝確率' + ('%g' % TAN_IDX_IRON_WP) + '%以上)。◎は連対確率1位の馬です。<span class="anav">橙色</span>は穴候補の馬。</dd>'
        + ('<dt>○▲△</dt><dd>◎の相手(紐)3頭。勝負レース（◎の単指数90以上）は単勝確率の高い順、それ以外のレースは期待値（勝率×予想オッズ）の高い順に選びます。</dd>' if REF_NEW_LOGIC else '') +
        '<dt>紐馬</dt><dd>紐馬指数(0〜100)。2〜3着に来る確率に、予想オッズの平方根(配当の大きさ)を掛けたもの'
        '(そのレースで最も紐向きの馬が100)。</dd>'
        '<dt>騎手指数</dt><dd>騎手の評価を偏差値で表したもの(50が平均)。その日に乗る他の馬の人気・騎乗数・減量騎手・'
        'この馬とのコンビ成績から計算。60以上は好騎手。</dd>'
        '<dt>緑の太字</dt><dd>注目の値(期待値110以上・騎手指数60以上)。</dd>'
        + ('<dt>軸級 S〜D</dt><dd>◎が1着になる推定確率の区分。%s / D=それ未満。</dd>'
           '<dt>信頼度</dt><dd>◎が3着以内に入る推定確率(%%)。%s。</dd>'
           % (' / '.join('%s=%g%%以上' % (_g, _c) for _g, _c in JUDGE_GRADE_CUTS),
              '・'.join('「%s」%g以上' % (_lb, _mn) for _mn, _lb, *_ in JUDGE_CONF_TIERS if _mn >= 0))
           if USE_NEW_JUDGE else
           '<dt>信頼度</dt><dd>S〜Dは軸馬の信頼度(Sが最も高い)。数字は0〜100で、そのレースの買い目が的中しやすいかの目安です。</dd>') +
        '<dt>勝負レース</dt><dd>◎の単指数が90以上のレース（黄色の枠）。◎から単勝確率の高い順に3頭へ、馬連・ワイド・馬単を表示します。組番は馬番です。</dd>'
        '<dt>参考買い目</dt><dd>勝負レース以外のレースで、◎から期待値（勝率×予想オッズ）の高い3頭への馬連・馬単（馬単は◎が1着・相手が2着）。参考としてご覧ください。</dd>'
        '</dl></div>')

    title = '%s 地方競馬予想%s' % (NOTE_BRAND, (' ' + date_label) if date_label else '')
    out = ['<!DOCTYPE html><html lang="ja"><head><meta charset="utf-8">' + bado_version_meta() +
           '<meta name="viewport" content="width=device-width,initial-scale=1">'
           '<title>%s</title><style>%s</style></head><body><div class="wrap">' % (_h(title), _NOTE_CSS),
           '<header class="cover"><div class="brand">%s ／ 地方競馬予想(オッズ×統計)</div>' % _h(NOTE_BRAND),
           '<h1>%s</h1>' % _h(date_label or '地方競馬予想'),
           '<div class="venues">%s</div>' % _h('・'.join(str(v) for v in venues)),
           '<div class="kpis"><div class="kpi"><b>%d</b><span>予想レース</span></div>'
           '<div class="kpi"><b>%d</b><span>勝負レース</span></div>'
           '<div class="kpi"><b>%d</b><span>勝負レースの買い目（点）</span></div></div></header>'
           % (len(cards), n_pick_races, n_pick_pts),
           '<h2>本日の勝負レース（発走順）</h2>', index_html,
           '<h2>予想表の見方</h2>', guide_html,
           '<h2>全レースの予想</h2>']
    out.extend(cards)
    out.append('<div class="foot">%s<div class="ver">%s %s</div></div></div></body></html>'
               % (_h(NOTE_DISCLAIMER), _h(NOTE_BRAND), _h(BADO_VERSION)))
    return '\n'.join(out)


# ──────────────────────────────────────────────────────────────
# Excel / CSV 出力（v077完全版）
# ──────────────────────────────────────────────────────────────
def _ref_bet_lines(analysis):
    """勝負レース/参考の買い目(◎→○▲△)を [(券種名, [組番...], ラベル, 印の並び)] で返す。"""
    out = []
    try:
        _tb = analysis['table']
        _mk = {int(b): str(m or '') for b, m in zip(_tb['番'], _tb['印'])}
    except Exception:
        _mk = {}
    for rf in analysis.get('reference_bets') or []:
        kn = GZ_KIND_JP.get(rf.get('kind'), rf.get('kind'))
        sep = GZ_KIND_SEP.get(rf.get('kind'), '-')
        combos = ['%d%s%d' % (rf['axis_ban'], sep, b) for b in rf.get('partner_bans', [])]
        marks = [_mk.get(int(b), '') for b in rf.get('partner_bans', [])]
        lab = ('%s｜◎から単勝確率の高い3頭' % SHOBU_LABEL) if rf.get('shobu') else str(rf.get('desc') or '参考')
        out.append((kn, combos, lab, marks))
    return out


def _write_bet_table(ws, current_row, title, headers, rows,
                     title_fill_color=None, row_fill_func=None):
    """表形式で買い目を書き込むヘルパー"""
    _header_fill = PatternFill(start_color="2E75B6", end_color="2E75B6", fill_type="solid")
    _header_font = Font(bold=True, color="FFFFFF", size=11)
    _thin_border = Border(
        left=Side(style='thin'), right=Side(style='thin'),
        top=Side(style='thin'), bottom=Side(style='thin')
    )
    # タイトル
    title_cell = ws.cell(row=current_row, column=1, value=title)
    title_cell.font = Font(bold=True, color="C00000" if "単勝" in title or "複勝" in title else "000000")
    if title_fill_color:
        title_cell.fill = title_fill_color
    current_row += 1

    # ヘッダー行
    for col_idx, h in enumerate(headers, 1):
        c = ws.cell(row=current_row, column=col_idx, value=h)
        c.font = _header_font
        c.fill = _header_fill
        c.border = _thin_border
        c.alignment = Alignment(horizontal='center')
    current_row += 1

    # データ行
    for row_data in rows:
        for col_idx, v in enumerate(row_data, 1):
            c = ws.cell(row=current_row, column=col_idx, value=v)
            c.border = _thin_border
            c.alignment = Alignment(horizontal='center')
            if row_fill_func:
                row_fill_func(c, row_data)
        current_row += 1
    current_row += 1  # セクション間スペース
    return current_row


def _run_output(grouped, output_csv, excel_path):
    """CSV + Excel 出力処理（v077完全版）"""
    _HIGH_YELLOW_HEX = "FFEB9C"

    def _fill_ev110(c, row_data):
        """期待値(index=3)が90以上のセルを黄色に塗る"""
        if isinstance(row_data[3], (int, float)) and row_data[3] >= 90:
            c.fill = PatternFill(start_color=_HIGH_YELLOW_HEX, end_color=_HIGH_YELLOW_HEX, fill_type="solid")

    try:
        with open(output_csv, 'w', encoding='utf-8-sig') as f:
            wb = Workbook()
            ws = wb.active
            ws.title = "地方競馬予想レポート"

            # スタイル定義
            title_fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
            title_font = Font(bold=True, color="FFFFFF", size=14)
            good_green = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
            high_yellow = PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid")
            hole_orange = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")
            iron_red = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
            header_fill = PatternFill(start_color="2E75B6", end_color="2E75B6", fill_type="solid")
            header_font = Font(bold=True, color="FFFFFF", size=11)
            thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
            high_single_fill = PatternFill(start_color="D5E8D4", end_color="D5E8D4", fill_type="solid")
            low_odds_fill = PatternFill(start_color="DAE8FC", end_color="DAE8FC", fill_type="solid")
            new_s_axis_fill = PatternFill(start_color="B4C6E7", end_color="B4C6E7", fill_type="solid")
            himba_idx_fill = PatternFill(start_color="F4B183", end_color="F4B183", fill_type="solid")

            current_row = 1
            ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=14)
            ws.cell(row=current_row, column=1, value="BADO 地方競馬予想 v137_001 オッズ×統計 アンサンブル")
            ws.cell(row=current_row, column=1).fill = title_fill
            ws.cell(row=current_row, column=1).font = title_font
            ws.cell(row=current_row, column=1).alignment = Alignment(horizontal='center')
            current_row += 1

            ws.cell(row=current_row, column=1, value=" 単勝確率=オッズ勝率×統計勝率の融合 | 単勝オッズ=予想オッズ(前売×人気オッズ) | 前売オッズ=補正前 | 馬番【緑】=単指数90以上(鉄板クラス) | 馬名【黄】=単指数90以上(鉄板馬) | 淡い緑=単勝確率が前売の評価より15%以上高い | 淡い青=予想オッズ9.9以下 | 淡い藍=騎手指数60以上 | 淡い橙=紐馬指数84以上 | 黄=期待値100以上 ")
            current_row += 2

            if grouped.ngroups == 0:
                ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=14)
                ws.cell(row=current_row, column=1, value="【注意】レースデータが読み込めませんでした（DATA_PATHまたはCSV内容を確認）。v063統計モデルは正常に動作しています。サンプルとしてこのメッセージを表示。")
                ws.cell(row=current_row, column=1).fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
                current_row += 2

            _first_race_error_shown = [False]
            RUN_ERRORS.clear()   # ★v136_021
            _html_races = []   # ★HTML予想表用: (name, analysis, extras)
            for name, group in grouped:
                try:
                    analysis = generate_analysis(group.copy())
                    if analysis['table'].empty:
                        f.write(f'No horses for group {name}\n')
                        continue

                    # ★HTML用: 馬番→(枠/性齢/斤量/騎手指数) を元groupから収集
                    # ★v135_006 fix: 騎手補正(KISHU_DEV_COL)は generate_analysis の
                    #   内部で追加される列であり、呼び出し側の group には存在しない。
                    #   そのため analysis['table'] から 番→値 を引く。
                    _kdev_map = {}
                    _atbl = analysis.get('table')
                    if _atbl is not None and not _atbl.empty \
                            and KISHU_DEV_COL in _atbl.columns:
                        for _, _tr in _atbl.iterrows():
                            _tb = _to_int_z2h(_tr.get('番'))
                            if _tb is not None:
                                _kdev_map[_tb] = _tr.get(KISHU_DEV_COL, '')
                    _extras = {}
                    # 騎手変更(horselist)で差し替えた騎手指数。group は元CSVのままなので上書きする。
                    _jk_new_idx = {c['ban']: c['idx_new'] for c in (analysis.get('dq_jk') or [])
                                   if c.get('idx_new') is not None}
                    for _, _gr in group.iterrows():
                        _b = _to_int_z2h(_gr.get('番'))
                        if _b is None:
                            continue
                        _kin = pd.to_numeric(_gr.get('斤量'), errors='coerce') if '斤量' in group.columns else np.nan
                        _extras[_b] = dict(
                            waku=_to_int_z2h(_gr.get('枠')) if '枠' in group.columns else None,
                            seirei=str(_gr.get('性齢', '') or ''),
                            kinryo=('' if _kin != _kin else ('%g' % _kin)),
                            kishu_idx=(int(float(_jk_new_idx.get(_b, _gr.get('騎手指数', 0)) or 0))
                                       if '騎手指数' in group.columns else ''),
                            # ★v135_006: 表示専用の補正騎手指数(偏差値)
                            #   analysis['table'] 由来のマップから取得する
                            kishu_dev=_kdev_map.get(_b, ''),
                        )
                    _html_races.append((name, analysis, _extras))

                    # 軸馬情報と鉄板馬フラグを事前計算
                    axis_ban_str = '未定'
                    axis_name_str = ''
                    if not analysis['table'].empty:
                        axis_row_local = analysis['table'][analysis['table']['印'] == '◎']
                        if not axis_row_local.empty:
                            axis_ban_str = str(int(axis_row_local['番'].iloc[0]))
                            axis_name_str = axis_row_local['馬名'].iloc[0]
                    is_iron_horse = _axis_is_iron(analysis)

                    # ★v137_001: 鉄板馬 = 単指数90以上(単勝確率 TAN_IDX_IRON_WP 以上)
                    iron_mask = (analysis['table']['単指数'] >= TAN_IDX_IRON)
                    has_iron = bool(iron_mask.any())
                    iron_bans = set(int(b) for b in analysis['table'].loc[iron_mask, '番'].dropna().astype(int).tolist()) if has_iron else set()

                    # ========== Excel書き込み ==========
                    race_title = f"{name[0]} {_dist_txt(name[1])} {name[2]}R ({name[3] or '-'}) | 荒れ度:{analysis['payout_level_score']} | {analysis['arare_class']}"
                    _is_shobu = bool(analysis.get('shobu'))
                    if _is_shobu:
                        race_title = '★%s(単指数%.1f) ' % (SHOBU_LABEL, float(analysis.get('shobu_idx') or 0)) + race_title
                    ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=14)
                    cell = ws.cell(row=current_row, column=1, value=race_title)
                    cell.fill = PatternFill(start_color="C00000", end_color="C00000", fill_type="solid") if _is_shobu else header_fill
                    cell.font = header_font
                    current_row += 1

                    headers = ['番','馬名','騎手',KISHU_DEV_LABEL,'展開','前売オッズ','人気オッズ','単指数','紐馬指数','印','単勝人気','単勝オッズ','単勝確率','複勝確率','単勝期待値','統計印','統計勝率','オッズ勝率']
                    for col_idx, h in enumerate(headers, 1):
                        c = ws.cell(row=current_row, column=col_idx, value=h)
                        c.fill = header_fill
                        c.font = header_font
                        c.border = thin_border
                        c.alignment = Alignment(horizontal='center')
                    current_row += 1

                    def _nv(v, nd=1):
                        try:
                            v = float(v)
                            return round(v, nd) if v == v else ''
                        except Exception:
                            return ''
                    for _, row in analysis['table'].iterrows():
                        pop = int(row['単勝人気']) if not pd.isna(row['単勝人気']) else 99
                        fuku_prob = round(row['複勝確率'], 1)
                        ban_no = int(row['番']) if not pd.isna(row['番']) else None
                        is_green_axis = (_nv(row['単指数']) or 0) >= TAN_IDX_IRON
                        is_iron_horse_row = ban_no in iron_bans if has_iron else False
                        try:
                            _val_hi = float(row['単勝確率']) >= 1.15 * float(row['オッズ勝率']) and float(row['単勝確率']) >= 5
                        except Exception:
                            _val_hi = False
                        vals = [
                            ban_no if ban_no is not None else '',
                            row['馬名'], row['騎手'],
                            _nv(row.get(KISHU_DEV_COL)),
                            str(row['展開']),
                            _nv(row['前売オッズ']), _nv(row['人気オッズ']),
                            _nv(row['単指数']), _nv(row['紐馬指数']), row['印'],
                            pop,
                            round(float(row['単勝オッズ']), 1),
                            round(float(row['単勝確率']), 1),
                            fuku_prob,
                            int(row['単勝期待値']) if not pd.isna(row['単勝期待値']) else '',
                            str(row.get('統計印', '') or ''),
                            _nv(row.get('統計勝率')),
                            _nv(row.get('オッズ勝率')),
                        ]
                        for col_idx, v in enumerate(vals, 1):
                            c = ws.cell(row=current_row, column=col_idx, value=v)
                            c.border = thin_border
                            c.alignment = Alignment(horizontal='center')
                            if col_idx == 1 and is_green_axis:
                                c.fill = PatternFill(start_color="00B050", end_color="00B050", fill_type="solid")
                                c.font = Font(bold=True, color="FFFFFF")
                            if col_idx == 2 and is_iron_horse_row:
                                c.fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
                                c.font = Font(bold=True, color="000000")
                            if col_idx == 4 and isinstance(v, (int, float)) and v >= 60:
                                c.fill = new_s_axis_fill
                            if col_idx == 6 and _val_hi:
                                c.fill = high_single_fill
                            if col_idx == 9 and isinstance(v, (int, float)) and v >= 84:
                                c.fill = himba_idx_fill
                            if col_idx == 10 and v == '◎':
                                c.fill = iron_red
                                c.font = Font(bold=True)
                            elif col_idx == 10 and '穴' in str(v):
                                c.fill = hole_orange
                                c.font = Font(bold=True, color="A04A00")
                            if col_idx == 11 and 1 <= pop <= 4:
                                c.fill = PatternFill(start_color="CFE2F3", end_color="CFE2F3", fill_type="solid")
                            if col_idx == 12 and isinstance(v, (int, float)) and v <= 9.9:
                                c.fill = low_odds_fill
                            if col_idx == 13 and isinstance(v, (int, float)) and v >= 44:
                                c.fill = good_green
                            if col_idx == 14 and fuku_prob >= 64:
                                c.fill = PatternFill(start_color="B6D7A8", end_color="B6D7A8", fill_type="solid")
                            if col_idx == 15 and isinstance(v, (int, float)) and v >= 100:
                                c.fill = high_yellow
                        current_row += 1

                    current_row += 1

                    # ========== 軸馬サマリー ==========
                    ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=6)
                    title_cell = ws.cell(row=current_row, column=1, value="【軸馬サマリー】")
                    title_cell.font = Font(bold=True, color="FFFFFF", size=12)
                    title_cell.fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
                    title_cell.alignment = Alignment(horizontal='center')
                    current_row += 1

                    axis_data = [
                        ["軸馬", f"{axis_ban_str}番 {axis_name_str}"],
                        ["分析値", f"{analysis['axis_analysis']}/100  (妙味{analysis['myomi_axis_score']}点 / z-score {analysis.get('axis_zscore', 0):.2f})"],
                        ["判定区分", f"{analysis.get('judgment_class', '')}"],
                        ["期待水準", f"{analysis.get('expected_win_rate', '')}  /  {analysis.get('expected_place_rate', '')}"],
                        ["コメント", analysis.get('judgment_comment', '')[:60] + ("..." if len(analysis.get('judgment_comment', '')) > 60 else "")]
                    ]

                    for col_idx, h in enumerate(["項目", "内容"], 1):
                        c = ws.cell(row=current_row, column=col_idx, value=h)
                        c.fill = header_fill
                        c.font = header_font
                        c.border = thin_border
                        c.alignment = Alignment(horizontal='center')
                    current_row += 1

                    for item, value in axis_data:
                        c1 = ws.cell(row=current_row, column=1, value=item)
                        c1.border = thin_border
                        c1.font = Font(bold=True)
                        c1.fill = PatternFill(start_color="D9E2F3", end_color="D9E2F3", fill_type="solid")
                        c1.alignment = Alignment(horizontal='center')

                        c2 = ws.cell(row=current_row, column=2, value=value)
                        c2.border = thin_border
                        c2.alignment = Alignment(horizontal='left', wrap_text=True)

                        # v116: 「分析値」内容欄の色マークは廃止。判定区分は軸区分(S/A/B/C/D級)で色分け(GUIと同配色)。
                        if item == "判定区分":
                            _jc = str(analysis.get("judgment_class", ""))
                            _gl = (_jc.split(' / ')[0].replace('級軸', '').strip() or 'D')[:1].upper()
                            _gmap = {'S': 'FFD700', 'A': 'FF7A45', 'B': '58A6FF', 'C': '3FB950', 'D': '8B949E'}
                            _hex = _gmap.get(_gl, '8B949E')
                            c2.fill = PatternFill(start_color=_hex, end_color=_hex, fill_type="solid")
                            c2.font = Font(bold=True, color=("000000" if _gl == 'S' else "FFFFFF"))

                        current_row += 1

                    current_row += 1

                    # ========== 的中信頼度ブロック ==========
                    rec_result = calculate_confidence_score(analysis)
                    rec_score    = rec_result['score']
                    rec_category = rec_result['category']
                    rec_bg_hex   = rec_result['bg_hex']
                    rec_font_hex = rec_result['font_hex']
                    rec_bar_hex  = rec_result['bar_hex']
                    rec_breakdown= rec_result['breakdown']

                    ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=6)
                    rec_title_cell = ws.cell(row=current_row, column=1, value="【的中信頼度】")
                    rec_title_cell.font = Font(bold=True, color="FFFFFF", size=12)
                    rec_title_cell.fill = PatternFill(start_color="4A4A4A", end_color="4A4A4A", fill_type="solid")
                    rec_title_cell.alignment = Alignment(horizontal='center')
                    current_row += 1

                    c_label = ws.cell(row=current_row, column=1, value="信頼度スコア")
                    c_label.border = thin_border
                    c_label.font = Font(bold=True)
                    c_label.fill = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")
                    c_label.alignment = Alignment(horizontal='center')
                    score_text = f"{rec_score} / 100 点"
                    c_score = ws.cell(row=current_row, column=2, value=score_text)
                    c_score.border = thin_border
                    c_score.font = Font(bold=True, size=13)
                    c_score.alignment = Alignment(horizontal='center')
                    current_row += 1

                    c_label2 = ws.cell(row=current_row, column=1, value="信頼度ランク")
                    c_label2.border = thin_border
                    c_label2.font = Font(bold=True)
                    c_label2.fill = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")
                    c_label2.alignment = Alignment(horizontal='center')
                    c_cat = ws.cell(row=current_row, column=2, value=rec_category)
                    c_cat.border = thin_border
                    c_cat.fill = PatternFill(start_color=rec_bg_hex, end_color=rec_bg_hex, fill_type="solid")
                    c_cat.font = Font(bold=True, color=rec_font_hex, size=12)
                    c_cat.alignment = Alignment(horizontal='center')
                    current_row += 1

                    bar_filled = '■' * (rec_score // 10)
                    bar_empty  = '□' * (10 - rec_score // 10)
                    bar_text   = f"{bar_filled}{bar_empty}  {rec_score}点"
                    c_bar_label = ws.cell(row=current_row, column=1, value="スコアバー")
                    c_bar_label.border = thin_border
                    c_bar_label.font = Font(bold=True)
                    c_bar_label.fill = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")
                    c_bar_label.alignment = Alignment(horizontal='center')
                    ws.merge_cells(start_row=current_row, start_column=2, end_row=current_row, end_column=6)
                    c_bar = ws.cell(row=current_row, column=2, value=bar_text)
                    c_bar.border = thin_border
                    c_bar.fill = PatternFill(start_color=rec_bar_hex, end_color=rec_bar_hex, fill_type="solid")
                    c_bar.font = Font(bold=True, size=11)
                    c_bar.alignment = Alignment(horizontal='left')
                    current_row += 1

                    c_brk_label = ws.cell(row=current_row, column=1, value="算出内訳")
                    c_brk_label.border = thin_border
                    c_brk_label.font = Font(bold=True)
                    c_brk_label.fill = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")
                    c_brk_label.alignment = Alignment(horizontal='center')
                    ws.merge_cells(start_row=current_row, start_column=2, end_row=current_row, end_column=8)
                    c_brk = ws.cell(row=current_row, column=2, value=rec_breakdown)
                    c_brk.border = thin_border
                    c_brk.font = Font(size=9, color="444444")
                    c_brk.alignment = Alignment(horizontal='left', wrap_text=True)
                    ws.row_dimensions[current_row].height = 22
                    current_row += 2

                    is_high_payout    = analysis.get('popular_out_high_payout', False)
                    high_payout_score = analysis.get('pop_conf_score', 0.0)

                    _force_ref = analysis.get('force_reference', False)  # v132: 赤字開催地は全備考を参考に

                    if analysis.get('tan_recs'):
                        tan_rows = []
                        _yt = analysis.get('yuuryoku_tan', set())
                        _tpb = set(analysis.get('tan_priority_bans', []))
                        for rec in analysis['tan_recs']:
                            if rec.ban in _tpb:
                                suffix = '有力'   # 有力条件はゲートに優先
                            else:
                                suffix = '参考' if _force_ref else ('有力' if rec.ban in _yt else '参考')
                            tan_rows.append([rec.ban, rec.hit_rate, rec.odds, rec.ev, suffix])
                        current_row = _write_bet_table(ws, current_row, "単勝",
                            ["馬番", "的中率(%)", "オッズ", "期待値", "備考"], tan_rows, row_fill_func=_fill_ev110)

                    # ★v134: 安定運用(トリガミ覚悟)の単勝/複勝
                    _st_tan = analysis.get('stable_tan_rec')
                    if _st_tan is not None:
                        _tag = analysis.get('stable_tan_tag', '')
                        srows = [[_st_tan.ban, _st_tan.hit_rate, _st_tan.odds,
                                  _st_tan.ev, _tag]]
                        current_row = _write_bet_table(ws, current_row,
                            "単勝（安定運用）",
                            ["馬番", "的中率(%)", "オッズ", "期待値", "備考"], srows,
                            row_fill_func=_fill_ev110)
                    _st_fuku = analysis.get('stable_fuku_rec')
                    if _st_fuku is not None:
                        _tag = analysis.get('stable_fuku_tag', '')
                        srows = [[_st_fuku.ban, _st_fuku.hit_rate, _st_fuku.odds,
                                  _st_fuku.ev, _tag]]
                        current_row = _write_bet_table(ws, current_row,
                            "複勝（安定運用）",
                            ["馬番", "的中率(%)", "オッズ", "期待値", "備考"], srows,
                            row_fill_func=_fill_ev110)

                    # ★採用条件: 条件名の表示ラベル
                    _GZ_LABEL = {k: v['label']            # ★v135_003
                                 for k, v in GZ_COND_DEFS.items()}
                    _gz_ut_cond = analysis.get('gz_umatan_cond', {})
                    _gz_um_cond = analysis.get('gz_umaren_cond', {})
                    _gz_wd_cond = analysis.get('gz_wide_cond', {})   # ★v136_001
                    _gz_stakes  = analysis.get('gz_stakes', {})

                    def _conn_rows_gz(main_key, cond_map, kind, arrow, tier=None):
                        """★v135_011: tier を指定すると本線枠/紐カバー枠だけを抽出。"""
                        rows = []
                        for rec in analysis.get(main_key, []):
                            _k = (rec.ban1, rec.ban2)
                            if tier is not None and gz_tier(cond_map.get(_k, '')) != tier:
                                continue
                            _clabel = _GZ_LABEL.get(cond_map.get(_k, ''), '—')
                            _stake = _gz_stakes.get((rec.ban1, rec.ban2, kind), 0)
                            if _force_ref:
                                _bikou = '参考(見送り)'; _stake = 0
                            elif _stake <= 0:
                                _bikou = '見送り(点数上限)'
                            else:
                                _bikou = '★採用'
                            _stake_txt = '%d円' % _stake if _stake > 0 else '—'
                            if _stat_ura_of(analysis, cond_map.get(_k, '')):   # ★v136_020
                                _bikou += ' ' + STAT_URA_LABEL
                            rows.append([f"{rec.ban1}{arrow}{rec.ban2}", _clabel, rec.hit_rate,
                                         rec.odds, rec.ev, _stake_txt, _bikou])
                        return rows

                    # 条件別の色分け(採用条件を視覚的に区別)
                    _GZ_FILL = {k: PatternFill(start_color=v['fill'], end_color=v['fill'],
                                               fill_type="solid")
                                for k, v in GZ_COND_DEFS.items()}
                    def _fill_gz(cond_map, arrow):
                        def _f(cell, row_data):
                            # row_data[0]='ban1{arrow}ban2' から条件を引く
                            try:
                                a, b = row_data[0].split(arrow)
                                _k = (int(a), int(b))
                            except Exception:
                                return
                            fill = _GZ_FILL.get(cond_map.get(_k, ''))
                            if fill is not None:
                                cell.fill = fill
                        return _f

                    _gz_headers = ["軸馬→相手馬", "採用条件", "的中率(%)", "オッズ",
                                   "期待値", "賭け金(円)", "備考"]

                    # ★採用条件の見出しバナー(強調)
                    _gz_active = analysis.get('gz_active_conds', [])
                    if (analysis.get('umatan_recs') or analysis.get('umaren_recs')
                            or analysis.get('wide_recs')):
                        _banner = ws.cell(row=current_row, column=1,
                            value="◆◆ 採用条件 買い目（均等買い・1点固定・本線枠/紐カバー枠は別上限）◆◆")
                        _banner.font = Font(bold=True, color="FFFFFF", size=12)
                        _banner.fill = PatternFill(start_color="C00000", end_color="C00000", fill_type="solid")
                        for _cc in range(2, 8):
                            ws.cell(row=current_row, column=_cc).fill = PatternFill(
                                start_color="C00000", end_color="C00000", fill_type="solid")
                        current_row += 1
                        for _tk in ('main', 'cover'):
                            _acts = [c for c in _gz_active if gz_tier(c) == _tk]
                            if not _acts:
                                continue
                            _act_txt = ("発火条件【%s】: " % GZ_TIER_LABEL[_tk]
                                        + " / ".join(_GZ_LABEL.get(c, c) for c in _acts))
                            _actc = ws.cell(row=current_row, column=1, value=_act_txt)
                            _actc.font = Font(bold=True, size=10,
                                              color=("C00000" if _tk == 'main' else "7F6000"))
                            current_row += 1

                    # ★v135_011: 本線枠 → 紐カバー枠 の順に、枠ごとに表を分けて出力する。
                    #   紐カバー枠は「的中率が低い/下限が100%未満」の参考枠であることを
                    #   見出しに明記し、本線枠の買い目と混ざらないようにする。
                    for _tk, _tnote in (('main', '本線枠｜v137 新条件(オッズ×統計・未検証)・馬連2/ワイド1'),
                                        ('cover', '紐カバー枠｜本線と同一軸で相手だけ拡張。紐抜け回収用')):
                        # ★v136_001: ワイド → 馬連 → 馬単 の順(的中率の高い券種を上に)。
                        _wd_rows = _conn_rows_gz('wide_recs', _gz_wd_cond, 'wide', '-', _tk)
                        if _wd_rows:
                            current_row = _write_bet_table(
                                ws, current_row, "ワイド【%s】" % _tnote,
                                _gz_headers, _wd_rows, row_fill_func=_fill_gz(_gz_wd_cond, '-'))
                        _um_rows = _conn_rows_gz('umaren_recs', _gz_um_cond, 'umaren', '-', _tk)
                        if _um_rows:
                            current_row = _write_bet_table(
                                ws, current_row, "馬連【%s】" % _tnote,
                                _gz_headers, _um_rows, row_fill_func=_fill_gz(_gz_um_cond, '-'))
                        _ut_rows = _conn_rows_gz('umatan_recs', _gz_ut_cond, 'umatan', '→', _tk)
                        if _ut_rows:
                            current_row = _write_bet_table(
                                ws, current_row, "馬単【%s】" % _tnote,
                                _gz_headers, _ut_rows, row_fill_func=_fill_gz(_gz_ut_cond, '→'))

                    # ★v137_001: 勝負レース/参考の買い目(◎→○▲△・表示のみ)
                    for _kn, _combos, _lab, _mks in _ref_bet_lines(analysis):
                        current_row = _write_bet_table(
                            ws, current_row, "%s【%s】" % (_kn, _lab), ["軸馬→相手馬", "印"],
                            [[_c, _m] for _c, _m in zip(_combos, _mks)],
                            title_fill_color=(PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
                                              if _is_shobu else None))

                    # ★v136_020: 統計裏付け・観察(統計紐)の行(表示・記録のみ)
                    for _sl in _dq_text_lines(analysis) + _stat_text_lines(analysis):   # ★v136_021
                        _sc = ws.cell(row=current_row, column=1, value=_sl)
                        _sc.font = Font(bold=True, size=10, color="5B3F8F")
                        current_row += 1

                    current_row += 1

                    # ========== CSV出力 ==========
                    title = f"{name[0]}  {_dist_txt(name[1])} {name[2]}R ({name[3] or '-'}),,,,,,,,,,,, "
                    f.write(title + '\n')

                    f.write('馬番,馬名,騎手,' + KISHU_DEV_LABEL + ',展開,前売オッズ,人気オッズ,単指数,紐馬指数,印,単勝人気,単勝オッズ,単勝確率,複勝確率,単勝期待値,統計印,統計勝率,オッズ勝率\n')

                    unfold_map = {'逃': '逃げ', '先': '先行', '差': '差し', '追': '追込', '中': '中団', '捲': '捲り', 'マ': 'マーク'}
                    for _, row in analysis['table'].iterrows():
                        unfold_str = str(row['展開'])
                        for key, val in unfold_map.items():
                            if key in unfold_str:
                                unfold_str = unfold_str.replace(key, val)
                        def _c1(v):
                            try:
                                v = float(v)
                                return '%.1f' % v if v == v else ''
                            except Exception:
                                return ''
                        line = '{},{},{},{},{},{},{},{},{},{},{},{:.1f},{:.1f},{:.1f},{},{},{},{}\n'.format(
                            int(row['番']) if not pd.isna(row['番']) else '',
                            row['馬名'], row['騎手'],
                            _c1(row[KISHU_DEV_COL]) if KISHU_DEV_COL in row.index else '',
                            unfold_str,
                            _c1(row['前売オッズ']), _c1(row['人気オッズ']),
                            _c1(row['単指数']), _c1(row['紐馬指数']), row['印'],
                            int(row['単勝人気']) if not pd.isna(row['単勝人気']) else '',
                            float(row['単勝オッズ']),
                            float(row['単勝確率']), float(row['複勝確率']), row['単勝期待値'],
                            str(row.get('統計印', '') or ''),
                            _c1(row.get('統計勝率')), _c1(row.get('オッズ勝率')),
                        )
                        f.write(line)

                    f.write(',,,,,,,,,,,,,\n')
                    f.write('「レース分析」,,,,,,,,,,,,,\n')

                    f.write(f'軸馬: {axis_ban_str}番 {axis_name_str},,,,,,,,,,,,,\n')

                    is_iron_horse = _axis_is_iron(analysis)
                    axis_line = f'軸の分析値: {analysis["axis_analysis"]}/100  妙味軸度: {analysis["myomi_axis_score"]}点  軸z-score: {analysis.get("axis_zscore", 0):.2f}'
                    if is_iron_horse:
                        axis_line += ' ★単指数90以上(鉄板クラス)'
                    if analysis.get('high_confidence', False):
                        axis_line += ' 勝負高信頼'
                    f.write(axis_line + ',,,,,,,,,,,,,\n')

                    f.write(f'判定区分: {analysis["judgment_class"]},,,,,,,,,,,,,\n')
                    f.write(f'期待水準: {analysis["expected_win_rate"]} / {analysis["expected_place_rate"]} | {analysis["judgment_comment"]},,,,,,,,,,,,,\n')

                    f.write(',,,,,,,,,,,,,\n')
                    f.write(',,,,,,,,,,,,,\n')
                    f.write(',,,,,,,,,,,,,\n')

                    is_high_payout    = analysis.get('popular_out_high_payout', False)
                    high_payout_score = analysis.get('pop_conf_score', 0.0)

                    if is_high_payout:
                        f.write('【★高配当レース（的中率×高回収率 最高ロジック）】,,,,,,,,,,,,,\n')
                        f.write(f'高配当スコア: {high_payout_score}/100 （的中率と回収率の両立が期待できるレース）,,,,,,,,,,,,,\n')
                        f.write('優先推奨買い目（他の通常買い目は非表示）,,,,,,,,,,,,,\n')
                    else:
                        f.write('買い目(表示のみ):,,,,,,,,,,,,,\n')

                    f.write(',,,,,,,,,,,,,\n')
                    f.write('買い目,,,,,,,,,,,,,\n')

                    is_iron_horse = _axis_is_iron(analysis)
                    _force_ref = analysis.get('force_reference', False)  # v132

                    if analysis['tan_recs']:
                        f.write('単勝\n')
                        _yt = analysis.get('yuuryoku_tan', set())
                        _tpb = set(analysis.get('tan_priority_bans', []))
                        for rec in analysis['tan_recs']:
                            if rec.ban in _tpb:
                                suffix = ' 有力'
                            else:
                                suffix = ' 参考' if _force_ref else (' 有力' if rec.ban in _yt else ' 参考')
                            f.write(f'単勝 {rec.ban}, 的中率: {rec.hit_rate}%, オッズ: {rec.odds}, 期待値:, {rec.ev}{suffix}\n')
                        f.write('\n')

                    _gz_ut_cond_c = analysis.get('gz_umatan_cond', {})
                    _gz_um_cond_c = analysis.get('gz_umaren_cond', {})
                    _gz_wd_cond_c = analysis.get('gz_wide_cond', {})   # ★v136_001
                    _gz_stakes_c  = analysis.get('gz_stakes', {})
                    _GZ_LABEL_C = {                       # ★v135_003
                        k: '%s(%s %s)' % (k, {'umatan': '馬単', 'wide': 'ワイド'}.get(v['kind'], '馬連'),   # ★v136_021
                                          v['short'])
                        for k, v in GZ_COND_DEFS.items()}
                    if (analysis.get('umatan_recs') or analysis.get('umaren_recs')
                            or analysis.get('wide_recs')):   # ★v136_021 ワイドだけのレースも
                        f.write('★★ 採用条件 買い目（均等買い 1点固定 本線枠/紐カバー枠は別上限）★★\n')

                    def _write_gz_bets(recs, cond_map, kind, arrow, title, tier=None):
                        """馬連/馬単/ワイドの買い目行をテキスト出力(均等買い)。
                        ★v135_011: tier指定で本線枠/紐カバー枠を分けて出力する。"""
                        recs = [r for r in (recs or [])
                                if tier is None
                                or gz_tier(cond_map.get((r.ban1, r.ban2), '')) == tier]
                        if not recs:
                            return
                        f.write(title + '\n')
                        for rec in recs:
                            _cl = _GZ_LABEL_C.get(cond_map.get((rec.ban1, rec.ban2), ''), '—')
                            _st = _gz_stakes_c.get((rec.ban1, rec.ban2, kind), 0)
                            if _force_ref:
                                _bk = ' 参考(見送り)'; _st = 0
                            elif _st <= 0:
                                _bk = ' 見送り(点数上限)'
                            else:
                                _bk = ' ★採用'
                            _stxt = '%d円' % _st if _st > 0 else '—'
                            if _stat_ura_of(analysis, cond_map.get((rec.ban1, rec.ban2), '')):   # ★v136_020
                                _bk += ' ' + STAT_URA_LABEL
                            f.write(f'  {rec.ban1}{arrow}{rec.ban2}, {_cl}, 的中率: {rec.hit_rate}%, '
                                    f'オッズ: {rec.odds}, 期待値: {rec.ev}, 賭け金: {_stxt}{_bk}\n')
                        f.write('\n')

                    for _tk, _tnote in (('main', '本線枠'), ('cover', '紐カバー枠')):
                        # ★v136_001: ワイド → 馬連 → 馬単 の順。
                        _write_gz_bets(analysis.get('wide_recs'), _gz_wd_cond_c,
                                       'wide', '-', 'ワイド【%s】' % _tnote, _tk)
                        _write_gz_bets(analysis.get('umaren_recs'), _gz_um_cond_c,
                                       'umaren', '-', '馬連【%s】' % _tnote, _tk)
                        _write_gz_bets(analysis.get('umatan_recs'), _gz_ut_cond_c,
                                       'umatan', '→', '馬単【%s】' % _tnote, _tk)

                    for _kn, _combos, _lab, _mks in _ref_bet_lines(analysis):   # ★v137_001
                        f.write('%s【%s】 %s\n' % (_kn, _lab, ' / '.join(_combos)))
                    if analysis.get('shobu'):
                        f.write('★%s(単指数%.1f・鉄板クラス)\n' % (SHOBU_LABEL, float(analysis.get('shobu_idx') or 0)))
                    for _dl in _dq_text_lines(analysis):   # ★v136_021 データ注意
                        f.write('⚠ ' + _dl + '\n')
                    # ★v136_019: 統計予想(参考・実弾外)
                    _stc = analysis.get('stat')
                    if _stc:
                        f.write('統計予想（参考・実弾外）\n')
                        f.write('  統計印: ' + ' '.join('%s%d(%.1f%%)' % (_stc['marks'][_b], _b, _stc['p_fused'][_b])
                                                     for _b in _stc['order'][:4]) + '\n')
                        for _x in _stc['bets']:
                            f.write('  %s %d-%d, 推定的中確率: %.1f%%\n'
                                    % ('馬連' if _x['kind'] == 'umaren' else 'ワイド', _x['ban1'], _x['ban2'], _x['prob']))
                        for _sl in _stat_text_lines(analysis):   # ★v136_020
                            f.write('  ' + _sl + '\n')
                        f.write('\n')

                    f.write(',,,,,,,,,,,,,\n')
                    f.write(',,,,,,,,,,,,,\n')

                except Exception as e:
                    import traceback as _tb
                    _msg = f'Error processing group {name}: {type(e).__name__} - {str(e)}'
                    f.write(_msg + '\n')
                    # ★ 追加: レース処理失敗をExcelにも赤行で明示(空レポート化の原因を可視化)
                    try:
                        ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=14)
                        _c = ws.cell(row=current_row, column=1,
                                     value=f'【エラー】{name} の処理に失敗: {type(e).__name__}: {e}')
                        _c.fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
                        _c.font = Font(bold=True, color="9C0006")
                        current_row += 2
                    except Exception:
                        pass
                    # ★ 最初の失敗はコンソールに完全なトレースバックを出す(モデル未検出等の根本原因特定用)
                    # ★v136_021: 全件記録し、予想表HTMLの先頭とGUIに出す(旧版は最初の1件だけ表示)
                    try:
                        RUN_ERRORS.append('%s %sR: %s: %s' % (name[0], _race_no_of(name[2]), type(e).__name__, e))
                    except Exception:
                        RUN_ERRORS.append(f'{name}: {type(e).__name__}: {e}')
                    print(f"[RACE ERROR] {_msg}")
                    _tb.print_exc()

        print(f"✅ 予想CSVを '{output_csv.name}' に出力しました。")
        # f.flush() は不要（withブロック終了時に自動でflush+closeされるため削除）

        # Excel保存
        try:
            # ★v135_005: 補正騎手指数の列追加により 14→15 列
            for col in range(1, 16):
                ws.column_dimensions[get_column_letter(col)].width = 14
            ws.column_dimensions['D'].width = 10   # 補正騎手指数は狭めで足りる
            ws.column_dimensions['B'].width = 20
            wb.save(excel_path)
            print(f"✅ カラフルExcelを '{excel_path.name}' に出力しました！（v079版）")
        except Exception as ex:
            print(f"Excel保存エラー: {str(ex)}")

        # ★v136_021: データ確認(結合率・騎手変更・取消・分析エラー)
        _dq_q = data_quality_summary([a for _, a, _ in _html_races], RUN_ERRORS)
        for _dl in data_quality_lines(_dq_q):
            print('⚠ ' + _dl)

        # ★ HTML予想表(中央競馬スタイル)の出力
        try:
            if _html_races:
                html_path = Path(str(output_csv)).with_suffix('.html')
                _html = build_html_nar(_html_races, Path(str(output_csv)).stem)
                _bn = data_quality_banner_html(_dq_q)          # ★v136_021
                if _bn:
                    _html = _html.replace('</head><body>', '</head><body>' + _bn, 1)
                with open(html_path, 'w', encoding='utf-8') as fh:
                    fh.write(_html)
                print(f"✅ HTML予想表を '{html_path.name}' に出力しました。")
                # ★v135_007: 同じ内容をPDFにも出力(失敗しても処理は継続)
                try:
                    pdf_path = html_path.with_suffix('.pdf')
                    _be = html_to_pdf(html_path, pdf_path)
                    if _be:
                        print(f"✅ PDF予想表を '{pdf_path.name}' に出力しました。"
                              f"(変換: {_be})")
                    else:
                        print("⚠ PDF変換をスキップしました(変換手段が見つかりません)。\n"
                              "   Windowsなら通常はEdgeで自動変換されます。"
                              "見つからない場合は次のいずれかを用意してください:\n"
                              "     ・Microsoft Edge / Google Chrome をインストール\n"
                              "     ・pip install playwright && playwright install chromium\n"
                              "     ・wkhtmltopdf を入れてPATHに追加\n"
                              f"   HTMLは出力済みなので、ブラウザで '{html_path.name}' を"
                              "開いて Ctrl+P → PDF保存でも同じ結果になります。")
                except Exception as ex:
                    print(f"PDF出力エラー(HTMLは出力済み): {str(ex)}")
        except Exception as ex:
            print(f"HTML出力エラー: {str(ex)}")

        # ★v136_020: 統計観察記録CSV
        try:
            if _html_races:
                _obs_path = Path(str(output_csv)).with_name(Path(str(output_csv)).stem + '_統計観察記録.csv')
                if write_stat_obs_log(_html_races, _obs_path):
                    print(f"✅ 統計観察記録を '{_obs_path.name}' に出力しました。")
        except Exception as ex:
            print(f"統計観察記録の出力エラー: {type(ex).__name__}: {ex}")

        # ★v136_007: note読者向けの予想表(<名前>_note.html / _note.pdf)
        #   研究用HTML(上)は keiba_tool が解析するので構造を変えず、別ファイルで出す。
        try:
            if _html_races and NOTE_HTML_ENABLE:
                _stem = Path(str(output_csv)).stem
                note_path = Path(str(output_csv)).with_name(_stem + '_note.html')
                _nhtml = build_html_note(_html_races, _stem)
                with open(note_path, 'w', encoding='utf-8') as fh:
                    fh.write(_nhtml)
                print(f"✅ 読者向け予想表を '{note_path.name}' に出力しました。")
                try:
                    note_pdf = note_path.with_suffix('.pdf')
                    _be = html_to_pdf(note_path, note_pdf, landscape=False)
                    if _be:
                        print(f"✅ 読者向けPDFを '{note_pdf.name}' に出力しました。(変換: {_be})")
                except Exception as ex:
                    print(f"読者向けPDF出力エラー(HTMLは出力済み): {str(ex)}")
        except Exception as ex:
            import traceback
            print(f"読者向けHTML出力エラー: {type(ex).__name__}: {ex}")
            traceback.print_exc()

    except Exception as e:
        import traceback
        print(f"[FATAL] 出力処理で致命的エラー: {type(e).__name__}: {e}")
        traceback.print_exc()
        # 最低限CSVファイルを作成してエラーログを残す
        try:
            with open(output_csv, 'w', encoding='utf-8-sig') as f_err:
                f_err.write(f"処理中に致命的エラーが発生しました: {e}\n")
                f_err.write("詳細はコンソールログを確認してください。\n")
        except:
            pass


# ══════════════════════════════════════════════════════════════
# ★★★ GUI 部分（tkinter）★★★
# ══════════════════════════════════════════════════════════════
# ══════════════════════════════════════════════════════════════
# ★v136_016: GUI デザイン刷新(予想ロジック・買い目・出力ファイルは v136_015 と同一)
#   ・配色は研究用HTML/note版と同じ『深緑×クリーム×金』。BADO_GUI_THEME=dark で暗色版。
#   ・左: データ選択 + レース一覧(級・信頼度・★採用を一覧表示、開催場/★採用で絞り込み)
#   ・右上: レース見出し(軸級・信頼度・レース分類・荒れ度をチップ表示)
#   ・中央: 出走馬テーブル(見出しクリックで並べ替え、◎・○▲△・穴を色分け)
#   ・下段: ◎軸馬カード(推定勝率/推定3着内率のメーター) + 推奨買い目(まとめ/券種別タブ)
#   ・高DPI(Windows)対応、フォントは BIZ UDPゴシック → Yu Gothic UI → メイリオ の順に自動選択。
#   ・ショートカット: Ctrl+O=CSVを開く / F5=予想実行・ファイル出力
# ══════════════════════════════════════════════════════════════
GUI_THEME = (_os.environ.get('BADO_GUI_THEME', 'light') or 'light').strip().lower()

_GUI_PALETTES = {
    'light': dict(
        BG='#EFECE3', SURFACE='#FFFFFF', SURFACE2='#F7F5EF', BORDER='#DCD6C8',
        BRAND='#13392B', BRAND2='#1D5541', BRAND_TEXT='#FFFFFF', GOLD='#C9A24A', GOLD2='#B38B34',
        TEXT='#1C2430', TEXT2='#4B5563', MUTED='#8A919B',
        SEL='#D7EBE1', SEL_TEXT='#0E2E22',
        AXIS_ROW='#FFF3D1', MARK_ROW='#EEF6F1', EVEN='#FFFFFF', ODD='#F8F7F3',
        ANA='#C85A12', GOOD='#127A53', WARN='#B7791F', BAD='#B42318',
        GZ='#B42318', GZ_BG='#FDECEC', SECTION_BG='#F1EEE5',
        HEAD_BG='#F1EEE5', HEAD_TEXT='#374151', TRACK='#E7E2D6',
    ),
    'dark': dict(
        BG='#0F1419', SURFACE='#171D24', SURFACE2='#1C242D', BORDER='#2A333D',
        BRAND='#123A2D', BRAND2='#1B5240', BRAND_TEXT='#F1F5F2', GOLD='#D4AF57', GOLD2='#BF9A43',
        TEXT='#E6EDF3', TEXT2='#B3BCC7', MUTED='#7D8792',
        SEL='#1F4A3A', SEL_TEXT='#F1F5F2',
        AXIS_ROW='#3A3016', MARK_ROW='#18302A', EVEN='#171D24', ODD='#1B2129',
        ANA='#F0A03A', GOOD='#3FB950', WARN='#D29922', BAD='#F47067',
        GZ='#FF7B72', GZ_BG='#3A1717', SECTION_BG='#202933',
        HEAD_BG='#1F2731', HEAD_TEXT='#C9D1D9', TRACK='#2A333D',
    ),
}
# 軸級バッジ (背景, 文字) … note版と同じ配色
_GUI_GRADE_STYLE = {'S': ('#F4C542', '#3B2A00'), 'A': ('#E0673A', '#FFFFFF'),
                    'B': ('#3B82C4', '#FFFFFF'), 'C': ('#3A9A5B', '#FFFFFF'),
                    'D': ('#8A929B', '#FFFFFF')}
# 信頼度区分の色
_GUI_TIER_COLOR = {'非常に高い': '#127A53', '高い': '#1F6FB2', '中程度': '#B7791F',
                   '低い': '#D9731A', '非常に低い': '#B42318'}
# レース分類の色
_GUI_CAT_COLOR = {'投資レース': '#8C2D19', '有力レース': '#1D5541',
                  '有力レース（2頭軸）': '#5B4B8A', '標準レース': '#6B7280', '見送りレース': '#6B7280'}


def run_gui():
    import tkinter as tk
    from tkinter import ttk, filedialog, messagebox
    import tkinter.font as tkfont

    # ── 高DPI(Windows)で文字がぼやけないようにする ──
    try:
        import ctypes
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass

    P = _GUI_PALETTES.get(GUI_THEME, _GUI_PALETTES['light'])
    BG, SURFACE, SURFACE2, BORDER = P['BG'], P['SURFACE'], P['SURFACE2'], P['BORDER']
    BRAND, BRAND2, BRAND_TEXT, GOLD, GOLD2 = P['BRAND'], P['BRAND2'], P['BRAND_TEXT'], P['GOLD'], P['GOLD2']
    TEXT, TEXT2, MUTED = P['TEXT'], P['TEXT2'], P['MUTED']

    # ── キャッシュ（レース名→analysis辞書） ────────────────────
    _cache: dict = {}

    app = tk.Tk()
    app.title(f'BADO 地方競馬予想  {BADO_VERSION}')
    app.configure(bg=BG)

    # ★v136_016a: 画面の拡大率(Windows の 125%/150% 等)に合わせてピクセル寸法を拡大する。
    #   フォントはポイント指定のため拡大率に応じて大きくなるが、列幅・パネル幅などの
    #   ピクセル指定はそのままだったため、高DPI環境で列が潰れて表示が崩れていた。
    try:
        S = max(1.0, float(app.winfo_fpixels('1i')) / 96.0)
    except Exception:
        S = 1.0

    # 画面が狭い(例: 1920x1080 を150%表示 → 実質1280x720)ときは、余白・パネル幅・文字を
    #   一律に詰める(D<1)。D は 0.72〜1.0。フォントは読みやすさのため 0.82 倍までに留める。
    _sw, _sh = app.winfo_screenwidth(), app.winfo_screenheight()
    D = max(0.72, min(1.0, (_sw / S) / 1500.0, (_sh / S - 48) / 900.0))
    DF = max(0.82, D)

    def px(n):
        return int(round(n * S * D))

    def fs(n):
        return max(8, int(n * DF + 0.5))

    _gw, _gh = min(px(1480), int(_sw * 0.94)), min(px(920), int(_sh * 0.88))
    app.geometry(f'{_gw}x{_gh}+{max(0, (_sw - _gw) // 2)}+{max(0, (_sh - _gh) // 3)}')
    app.minsize(min(px(1180), int(_sw * 0.9)), min(px(720), int(_sh * 0.8)))
    try:
        if D < 1.0:
            app.state('zoomed')          # 画面が小さいときは最大化(Windows)
    except Exception:
        pass

    # ── フォント(日本語が綺麗に出るものを自動選択) ──
    _fams = set(tkfont.families(app))
    _ff = next((f for f in ('BIZ UDPGothic', 'BIZ UDPゴシック', 'Yu Gothic UI', 'Meiryo UI',
                            'メイリオ', 'Meiryo', 'Hiragino Sans', 'Noto Sans CJK JP')
                if f in _fams), 'TkDefaultFont')
    F = {
        'brand':  tkfont.Font(family=_ff, size=fs(15), weight='bold'),
        'title':  tkfont.Font(family=_ff, size=fs(18), weight='bold'),
        'head':   tkfont.Font(family=_ff, size=fs(11), weight='bold'),
        'body':   tkfont.Font(family=_ff, size=fs(10)),
        'bodyb':  tkfont.Font(family=_ff, size=fs(10), weight='bold'),
        'small':  tkfont.Font(family=_ff, size=fs(9)),
        'smallb': tkfont.Font(family=_ff, size=fs(9), weight='bold'),
        'label':  tkfont.Font(family=_ff, size=fs(8), weight='bold'),
        'big':    tkfont.Font(family=_ff, size=fs(16), weight='bold'),
        'grade':  tkfont.Font(family=_ff, size=fs(20), weight='bold'),
        'num':    tkfont.Font(family=_ff, size=fs(13), weight='bold'),
    }
    _rowh = max(px(24), int(F['body'].metrics('linespace') * (1.7 if D >= 0.95 else 1.45)))

    def _colw(base, heading='', *samples):
        """列幅: 拡大率を掛けた基準幅と、見出し/代表値が収まる幅の大きい方。"""
        need = F['smallb'].measure(str(heading)) if heading else 0
        for _s in samples:
            need = max(need, F['bodyb'].measure(str(_s)))
        return max(px(base), need + px(12))

    # ── ttk スタイル ──────────────────────────────────────────
    style = ttk.Style(app)
    try:
        style.theme_use('clam')
    except Exception:
        pass
    style.configure('.', background=BG, foreground=TEXT, font=F['body'])
    style.configure('Horse.Treeview', background=SURFACE, fieldbackground=SURFACE, foreground=TEXT,
                    rowheight=_rowh, font=F['body'], borderwidth=0, relief='flat')
    style.configure('Horse.Treeview.Heading', background=P['HEAD_BG'], foreground=P['HEAD_TEXT'],
                    font=F['smallb'], relief='flat', borderwidth=0, padding=(px(4), px(6)))
    style.map('Horse.Treeview.Heading', background=[('active', P['SEL'])])
    style.map('Horse.Treeview', background=[('selected', P['SEL'])], foreground=[('selected', P['SEL_TEXT'])])
    style.layout('Horse.Treeview', [('Horse.Treeview.treearea', {'sticky': 'nswe'})])

    style.configure('Side.Treeview', background=SURFACE, fieldbackground=SURFACE, foreground=TEXT,
                    rowheight=_rowh, font=F['body'], borderwidth=0, relief='flat')
    style.configure('Side.Treeview.Heading', background=P['HEAD_BG'], foreground=P['HEAD_TEXT'],
                    font=F['smallb'], relief='flat', borderwidth=0, padding=(px(4), px(5)))
    style.map('Side.Treeview', background=[('selected', P['SEL'])], foreground=[('selected', P['SEL_TEXT'])])
    style.layout('Side.Treeview', [('Side.Treeview.treearea', {'sticky': 'nswe'})])

    for _orient in ('Vertical', 'Horizontal'):
        style.configure(f'Bado.{_orient}.TScrollbar', gripcount=0, background=BORDER,
                        troughcolor=SURFACE, bordercolor=SURFACE, lightcolor=BORDER,
                        darkcolor=BORDER, arrowcolor=MUTED, relief='flat')
        style.map(f'Bado.{_orient}.TScrollbar', background=[('active', MUTED)])

    style.configure('Bado.TNotebook', background=SURFACE, borderwidth=0, tabmargins=(px(8), px(6), px(8), 0),
                    bordercolor=BORDER, lightcolor=SURFACE, darkcolor=SURFACE)
    style.configure('Bado.TNotebook.Tab', background=SURFACE2, foreground=TEXT2, font=F['smallb'],
                    padding=(px(14), px(6)), borderwidth=0, bordercolor=BORDER, lightcolor=SURFACE2, darkcolor=SURFACE2)
    style.map('Bado.TNotebook.Tab', background=[('selected', SURFACE)],
              foreground=[('selected', BRAND if GUI_THEME != 'dark' else GOLD)],
              expand=[('selected', (0, 0, 0, 0))])

    style.configure('Bado.TCombobox', fieldbackground=SURFACE, background=SURFACE, foreground=TEXT,
                    arrowcolor=TEXT2, bordercolor=BORDER, lightcolor=BORDER, darkcolor=BORDER, padding=4)
    style.map('Bado.TCombobox', fieldbackground=[('readonly', SURFACE)], foreground=[('readonly', TEXT)],
              selectbackground=[('readonly', SURFACE)], selectforeground=[('readonly', TEXT)])
    app.option_add('*TCombobox*Listbox.background', SURFACE)
    app.option_add('*TCombobox*Listbox.foreground', TEXT)
    app.option_add('*TCombobox*Listbox.selectBackground', P['SEL'])
    app.option_add('*TCombobox*Listbox.selectForeground', P['SEL_TEXT'])
    app.option_add('*TCombobox*Listbox.font', F['body'])

    style.configure('Bado.TCheckbutton', background=SURFACE, foreground=TEXT2, font=F['small'])
    style.map('Bado.TCheckbutton', background=[('active', SURFACE)])
    style.configure('Bado.TPanedwindow', background=BG)
    style.configure('Sash', sashthickness=px(8), background=BG)

    # ── 共通部品 ──────────────────────────────────────────────
    def _card(parent, **kw):
        """白いカード(細い枠線付き)。内側の frame を返す。"""
        outer = tk.Frame(parent, bg=BORDER, bd=0, highlightthickness=0)
        inner = tk.Frame(outer, bg=SURFACE, bd=0, highlightthickness=0)
        inner.pack(fill='both', expand=True, padx=px(1), pady=px(1))
        outer.inner = inner
        return outer

    _BTN_KIND = {
        'primary':   (GOLD, '#1B1405', GOLD2),
        'brand':     (BRAND2, BRAND_TEXT, BRAND),
        'secondary': (SURFACE2, TEXT, BORDER),
        'ghost':     (SURFACE, BRAND if GUI_THEME != 'dark' else GOLD, SURFACE2),
    }

    def _button(parent, text, command, kind='secondary', font=None, padx=14, pady=7):
        """ホバーで色が変わるフラットボタン(tk.Label ベース)。.set_enabled(bool) で無効化。"""
        bg, fg, hover = _BTN_KIND[kind]
        b = tk.Label(parent, text=text, bg=bg, fg=fg, font=font or F['bodyb'],
                     padx=px(padx), pady=px(pady), cursor='hand2', bd=0)
        b._enabled = True

        def _enter(_e):
            if b._enabled:
                b.config(bg=hover)

        def _leave(_e):
            if b._enabled:
                b.config(bg=bg)

        def _click(_e):
            if b._enabled:
                command()

        def set_enabled(flag):
            b._enabled = bool(flag)
            b.config(bg=bg if flag else BORDER, fg=fg if flag else MUTED,
                     cursor='hand2' if flag else 'arrow')

        b.bind('<Enter>', _enter)
        b.bind('<Leave>', _leave)
        b.bind('<Button-1>', _click)
        b.set_enabled = set_enabled
        return b

    def _chip(parent, text, bg, fg='#FFFFFF', font=None):
        return tk.Label(parent, text=text, bg=bg, fg=fg, font=font or F['smallb'], padx=px(9), pady=px(3), bd=0)

    def _section_label(parent, text, bg=None):
        return tk.Label(parent, text=text, bg=bg or SURFACE, fg=MUTED, font=F['label'], anchor='w')

    def _fit_text(text, font, maxpx):
        """表示幅 maxpx に収まるよう中央を『…』で省略する(1行表示用)。"""
        s = str(text or '')
        if font.measure(s) <= maxpx:
            return s
        head, tail = s[:len(s) // 2], s[len(s) // 2:]
        while (head or tail) and font.measure(head + '…' + tail) > maxpx:
            if len(head) >= len(tail):
                head = head[:-1]
            else:
                tail = tail[1:]
        return head + '…' + tail

    def _abs_dir(pth):
        try:
            return str(Path(pth).expanduser().resolve())
        except Exception:
            return str(pth)

    def _ellipsis_path(p, n=40):
        s = str(p or '')
        return s if len(s) <= n else '…' + s[-(n - 1):]

    class _Meter(tk.Canvas):
        """横棒メーター(0-100)。"""
        def __init__(self, parent, color, height=10):
            super().__init__(parent, height=px(height), bg=SURFACE, highlightthickness=0, bd=0)
            self._v = 0.0
            self._color = color
            self.bind('<Configure>', lambda e: self._draw())

        def set(self, value, color=None):
            try:
                self._v = max(0.0, min(100.0, float(value)))
            except Exception:
                self._v = 0.0
            if color:
                self._color = color
            self._draw()

        def _draw(self):
            self.delete('all')
            w = max(self.winfo_width(), 2)
            h = max(self.winfo_height(), 2)
            r = h / 2.0

            def _pill(x0, x1, col):
                if x1 - x0 < h:
                    self.create_oval(x0, 0, x0 + h, h, fill=col, outline=col)
                    return
                self.create_oval(x0, 0, x0 + h, h, fill=col, outline=col)
                self.create_oval(x1 - h, 0, x1, h, fill=col, outline=col)
                self.create_rectangle(x0 + r, 0, x1 - r, h, fill=col, outline=col)

            _pill(0, w, P['TRACK'])
            if self._v > 0:
                _pill(0, max(h, w * self._v / 100.0), self._color)

    # ── 状態変数 ─────────────────────────────────────────────
    var_csv_path   = tk.StringVar(value='')
    var_out_dir    = tk.StringVar(value=str(Config.DEFAULT_OUTPUT_DIR))
    var_status     = tk.StringVar(value='単勝オッズCSV(nar_shutuba_all_*.csv)を開いてください(Ctrl+O)')
    var_progress   = tk.StringVar(value='')
    var_venue      = tk.StringVar(value='すべて')
    var_star_only  = tk.BooleanVar(value=False)
    var_race_count = tk.StringVar(value='')

    _current_analysis = [None]
    _current_race_name = [None]
    _grouped_data      = [None]
    _race_name_map     = {}   # iid → グループ名(タプル)
    _race_order        = []   # 読込順の iid
    _race_star         = {}   # iid → ★採用あり
    _load_token        = [0]  # CSV再読込で事前分析を打ち切るための世代番号
    _sort_state        = {'col': 'wp', 'asc': False}

    # ── ルート: ヘッダー / 本体 / ステータスバー ─────────────
    header = tk.Frame(app, bg=BRAND, height=px(58))
    header.pack(fill='x', side='top')
    header.pack_propagate(False)
    tk.Label(header, text='BADO', bg=BRAND, fg=GOLD, font=F['brand']).pack(side='left', padx=(px(20), px(8)))
    tk.Label(header, text='地方競馬予想', bg=BRAND, fg=BRAND_TEXT, font=F['brand']).pack(side='left')
    tk.Label(header, text=f' {BADO_VERSION} ', bg=BRAND2, fg=BRAND_TEXT, font=F['small'],
             padx=px(6), pady=px(2)).pack(side='left', padx=px(12))
    _logic_txt = '判定:%s ・ 参考予想:%s' % ('新' if USE_NEW_JUDGE else '旧',
                                        (('回収率重視' if REF_VALUE_LAMBDA >= 2 else '的中重視') + ' λ=%g' % REF_VALUE_LAMBDA) if REF_NEW_LOGIC else '旧')
    tk.Label(header, text=_logic_txt, bg=BRAND, fg='#A9C4B7', font=F['small']).pack(side='left')
    btn_run = _button(header, '▶  予想実行・ファイル出力  (F5)', lambda: _run_analysis(), kind='primary',
                      font=F['bodyb'], padx=18, pady=8)
    btn_run.pack(side='right', padx=px(20), pady=px(10))

    statusbar = tk.Frame(app, bg=SURFACE2, height=px(28), highlightthickness=1, highlightbackground=BORDER)
    statusbar.pack(fill='x', side='bottom')
    statusbar.pack_propagate(False)
    tk.Label(statusbar, textvariable=var_status, bg=SURFACE2, fg=TEXT2, font=F['small'],
             anchor='w').pack(side='left', padx=px(14))
    tk.Label(statusbar, textvariable=var_progress, bg=SURFACE2, fg=MUTED, font=F['small']).pack(side='right', padx=px(14))

    body = tk.Frame(app, bg=BG)
    body.pack(fill='both', expand=True)

    # ─── 左サイドバー ─────────────────────────────────────────
    sidebar = tk.Frame(body, bg=BG, width=px(340))
    sidebar.pack(side='left', fill='y', padx=(px(14), px(7)), pady=px(14))
    sidebar.pack_propagate(False)

    data_card = _card(sidebar)
    data_card.pack(fill='x')
    dc = data_card.inner
    _section_label(dc, 'データ').pack(fill='x', padx=px(14), pady=(px(12), px(2)))
    lbl_csv_name = tk.Label(dc, text='ファイル未選択', bg=SURFACE, fg=TEXT, font=F['bodyb'],
                            anchor='w', justify='left')
    lbl_csv_name.pack(fill='x', padx=px(14))
    lbl_csv_dir = tk.Label(dc, text='', bg=SURFACE, fg=MUTED, font=F['small'], anchor='w')
    lbl_csv_dir.pack(fill='x', padx=px(14), pady=(0, px(8)))
    btn_csv = _button(dc, 'オッズCSVを開く  (Ctrl+O)', lambda: _open_csv(), kind='brand')
    btn_csv.pack(fill='x', padx=px(14))
    # ★v136_019: 出馬表HTML(統計予想)
    btn_deba = _button(dc, '出馬表HTMLフォルダを読込', lambda: _open_deba(), kind='secondary')
    btn_deba.pack(fill='x', padx=px(14), pady=(px(6), 0))
    lbl_deba = tk.Label(dc, text=_fit_text('統計予想: ' + stat_summary(), F['small'], px(304)),
                        bg=SURFACE, fg=MUTED, font=F['small'], anchor='w', justify='left')
    lbl_deba.pack(fill='x', padx=px(14), pady=(px(2), 0))

    _od_row = tk.Frame(dc, bg=SURFACE)
    _od_row.pack(fill='x', padx=px(14), pady=(px(12), 0))
    _section_label(_od_row, '出力先フォルダ').pack(side='left')
    _button(_od_row, '変更', lambda: _choose_outdir(), kind='ghost', font=F['smallb'],
            padx=6, pady=0).pack(side='right')
    lbl_outdir = tk.Label(dc, text=_fit_text(_abs_dir(var_out_dir.get()), F['small'], px(300)), bg=SURFACE, fg=TEXT2,
                          font=F['small'], anchor='w')
    lbl_outdir.pack(fill='x', padx=px(14), pady=(px(2), px(10)))
    var_out_dir.trace_add('write', lambda *_: lbl_outdir.config(
        text=_fit_text(_abs_dir(var_out_dir.get()), F['small'], px(300))))

    btn_batch = _button(dc, 'フォルダ内を一括予想', lambda: _run_batch_prediction(
        app, var_status, var_out_dir, messagebox, filedialog, Path, pd), kind='secondary')
    btn_batch.pack(fill='x', padx=px(14), pady=(0, px(14)))

    race_card = _card(sidebar)
    race_card.pack(fill='both', expand=True, pady=(px(12), 0))
    rc = race_card.inner
    _rh = tk.Frame(rc, bg=SURFACE)
    _rh.pack(fill='x', padx=px(14), pady=(px(12), px(6)))
    _section_label(_rh, 'レース一覧').pack(side='left')
    tk.Label(_rh, textvariable=var_race_count, bg=SURFACE, fg=MUTED, font=F['small']).pack(side='right')

    _flt = tk.Frame(rc, bg=SURFACE)
    _flt.pack(fill='x', padx=px(14), pady=(0, px(8)))
    venue_cb = ttk.Combobox(_flt, textvariable=var_venue, values=['すべて'], state='readonly',
                            width=10, style='Bado.TCombobox', font=F['body'])
    venue_cb.pack(side='left')
    ttk.Checkbutton(_flt, text='★勝負のみ', variable=var_star_only, style='Bado.TCheckbutton',
                    command=lambda: _apply_filter()).pack(side='left', padx=(px(10), 0))
    venue_cb.bind('<<ComboboxSelected>>', lambda e: _apply_filter())

    _rl = tk.Frame(rc, bg=SURFACE)
    _rl.pack(fill='both', expand=True, padx=(px(8), px(4)), pady=(0, px(8)))
    race_tree = ttk.Treeview(_rl, style='Side.Treeview', selectmode='browse',
                             columns=('race', 'time', 'grade', 'conf', 'star'), show='headings')
    for _c, _h, _w, _a, _smp in (('race', 'レース', 110, 'w', '名古屋 12R'), ('time', '発走', 54, 'center', '20:30'),
                                 ('grade', '級', 34, 'center', 'S'), ('conf', '信頼', 44, 'center', '100'),
                                 ('star', '★', 30, 'center', '★')):
        race_tree.heading(_c, text=_h, anchor=_a)
        race_tree.column(_c, width=_colw(_w, _h, _smp), minwidth=_colw(_w, _h, _smp) if _c != 'race' else px(60),
                         anchor=_a, stretch=(_c == 'race'))
    race_sb = ttk.Scrollbar(_rl, orient='vertical', command=race_tree.yview, style='Bado.Vertical.TScrollbar')
    race_tree.configure(yscrollcommand=race_sb.set)
    race_sb.pack(side='right', fill='y')
    race_tree.pack(side='left', fill='both', expand=True)
    race_tree.tag_configure('even', background=P['EVEN'])
    race_tree.tag_configure('odd', background=P['ODD'])
    race_tree.tag_configure('gz', foreground=P['GZ'], font=F['bodyb'])
    race_tree.tag_configure('pending', foreground=MUTED)

    # ─── メインエリア ─────────────────────────────────────────
    main = tk.Frame(body, bg=BG)
    main.pack(side='left', fill='both', expand=True, padx=(px(7), px(14)), pady=px(14))

    # レース見出しカード
    head_card = _card(main)
    head_card.pack(fill='x')
    hc = head_card.inner
    _hc_top = tk.Frame(hc, bg=SURFACE)
    _hc_top.pack(fill='x', padx=px(18), pady=(px(12), 0))
    lbl_race_title = tk.Label(_hc_top, text='レースを選択してください', bg=SURFACE, fg=TEXT,
                              font=F['title'], anchor='w')
    lbl_race_title.pack(side='left')
    lbl_race_meta = tk.Label(_hc_top, text='', bg=SURFACE, fg=TEXT2, font=F['body'], anchor='w')
    lbl_race_meta.pack(side='left', padx=(px(14), 0), pady=(px(6), 0))
    chip_row = tk.Frame(hc, bg=SURFACE)
    chip_row.pack(fill='x', padx=px(18), pady=(px(8), px(12)))
    tk.Label(chip_row, text='左の一覧からレースを選ぶと、軸級・信頼度・買い目が表示されます。',
             bg=SURFACE, fg=MUTED, font=F['small']).pack(side='left')

    # 上下に分割(出走馬テーブル / 軸馬カード+買い目)
    paned = ttk.PanedWindow(main, orient='vertical', style='Bado.TPanedwindow')
    paned.pack(fill='both', expand=True, pady=(px(12), 0))

    # 出走馬テーブル
    table_card = _card(paned)
    paned.add(table_card, weight=3)
    tcard = table_card.inner
    _th = tk.Frame(tcard, bg=SURFACE)
    _th.pack(fill='x', padx=px(14), pady=(px(10), px(6)))
    _section_label(_th, '出走馬').pack(side='left')
    _lg = tk.Frame(_th, bg=SURFACE)
    _lg.pack(side='right')

    def _legend(txt, bg, fg=TEXT2):
        tk.Label(_lg, text=txt, bg=bg, fg=fg, font=F['small'], padx=px(6), pady=px(1)).pack(side='left', padx=(px(6), 0))
    _legend('◎ 本命(軸)', P['AXIS_ROW'])
    _legend('○▲△ 参考買い目の相手', P['MARK_ROW'])
    _legend('穴 穴条件', SURFACE, P['ANA'])
    tk.Label(_lg, text='・見出しクリックで並べ替え', bg=SURFACE, fg=MUTED, font=F['small']).pack(side='left', padx=(px(8), 0))

    TV = (  # (列ID, 見出し, 幅, 寄せ, 並べ替えキー列)
        ('mk',   '印',        46,  'center', None),   # (幅は _colw で見出し/代表値に合わせて拡大)
        ('ban',  '番',        40,  'center', '番'),
        ('name', '馬名',      150, 'w',      None),
        ('jk',   '騎手',      84,  'w',      None),
        ('kdev', KISHU_DEV_LABEL, 70, 'center', KISHU_DEV_COL),
        ('sty',  '脚質',      46,  'center', None),
        ('pop',  '人気',      46,  'center', '単勝人気'),
        ('odds', '単オッズ',  70,  'e',      '単勝オッズ'),
        ('wp',   '勝率%',     64,  'e',      '単勝確率'),
        ('fp',   '複勝%',     64,  'e',      '複勝確率'),
        ('ev',   '期待値',    60,  'e',      '単勝期待値'),
        ('ti',   '前売',      58,  'e',      '前売オッズ'),     # ★v137_001 補正前単勝オッズ
        ('fi',   '人気O',     58,  'e',      '人気オッズ'),     # ★v137_001 人気オッズ
        ('ens',  '単指数',  70,  'e',      '単指数'),
        ('himo', '紐馬指数',  70,  'e',      '紐馬指数'),
        ('smk',  '統計印',    52,  'center', None),        # ★v136_019
        ('sp',   '統計%',     60,  'e',      '統計勝率'),
        ('fu',   'オッズ%',   60,  'e',      'オッズ勝率'),
    )
    _tv_frame = tk.Frame(tcard, bg=SURFACE)
    _tv_frame.pack(fill='both', expand=True, padx=(px(8), px(4)), pady=(0, px(8)))
    tree_table = ttk.Treeview(_tv_frame, style='Horse.Treeview', columns=[c[0] for c in TV],
                              show='headings', selectmode='browse')
    for _cid, _hd, _w, _an, _key in TV:
        tree_table.heading(_cid, text=_hd, anchor=_an,
                           command=(lambda c=_cid: _on_heading(c)) if _key else '')
        _smp = {'mk': '◎穴', 'ban': '16', 'name': 'アークリオーソ', 'jk': '吉村智洋',
                'kdev': '62.7', 'sty': '差', 'pop': '16', 'odds': '155.0', 'wp': '39.6',
                'fp': '58.0', 'ev': '1234', 'ti': '100', 'fi': '25', 'ens': '33.7', 'himo': '94.0', 'smk': '◎', 'sp': '33.3', 'fu': '33.3'}.get(_cid, '')
        _cw = _colw(_w, _hd + ' ▼', _smp)
        tree_table.column(_cid, width=_cw, anchor=_an, minwidth=_cw if _cid != 'name' else px(100),
                          stretch=(_cid == 'name'))
    tsb_y = ttk.Scrollbar(_tv_frame, orient='vertical', command=tree_table.yview, style='Bado.Vertical.TScrollbar')
    tsb_x = ttk.Scrollbar(_tv_frame, orient='horizontal', command=tree_table.xview, style='Bado.Horizontal.TScrollbar')
    tree_table.configure(yscrollcommand=tsb_y.set, xscrollcommand=tsb_x.set)
    tsb_y.pack(side='right', fill='y')
    tsb_x.pack(side='bottom', fill='x')
    tree_table.pack(fill='both', expand=True)
    tree_table.tag_configure('even', background=P['EVEN'])
    tree_table.tag_configure('odd', background=P['ODD'])
    tree_table.tag_configure('mark', background=P['MARK_ROW'])
    tree_table.tag_configure('axis', background=P['AXIS_ROW'], font=F['bodyb'])
    tree_table.tag_configure('ana', foreground=P['ANA'])

    # 下段: 軸馬カード + 買い目カード
    lower = tk.Frame(paned, bg=BG)
    paned.add(lower, weight=2)

    axis_card = _card(lower)
    axis_card.pack(side='left', fill='y', padx=(0, px(12)))
    ac = axis_card.inner
    ac.config(width=px(380))
    ac.pack_propagate(False)
    axis_card.config(width=px(382))
    _section_label(ac, '◎ 軸馬').pack(fill='x', padx=px(16), pady=(px(12), px(4)))
    _ax_top = tk.Frame(ac, bg=SURFACE)
    _ax_top.pack(fill='x', padx=px(16))
    grade_badge = tk.Label(_ax_top, text='–', bg=BORDER, fg='#FFFFFF', font=F['grade'], width=2, pady=px(2))
    grade_badge.pack(side='left')
    _ax_names = tk.Frame(_ax_top, bg=SURFACE)
    _ax_names.pack(side='left', fill='x', expand=True, padx=(px(12), 0))
    lbl_axis_name = tk.Label(_ax_names, text='─', bg=SURFACE, fg=TEXT, font=F['big'], anchor='w')
    lbl_axis_name.pack(fill='x')
    lbl_grade_desc = tk.Label(_ax_names, text='', bg=SURFACE, fg=TEXT2, font=F['small'], anchor='w')
    lbl_grade_desc.pack(fill='x')

    def _meter_row(parent, title):
        fr = tk.Frame(parent, bg=SURFACE)
        fr.pack(fill='x', padx=px(16), pady=(px(10), 0))
        top = tk.Frame(fr, bg=SURFACE)
        top.pack(fill='x')
        tk.Label(top, text=title, bg=SURFACE, fg=TEXT2, font=F['small']).pack(side='left')
        val = tk.Label(top, text='–', bg=SURFACE, fg=TEXT, font=F['num'])
        val.pack(side='right')
        tag = tk.Label(top, text='', bg=SURFACE, fg=TEXT2, font=F['smallb'])
        tag.pack(side='right', padx=(0, px(8)))
        m = _Meter(fr, BRAND2)
        m.pack(fill='x', pady=(px(3), 0))
        return val, tag, m

    lbl_win_val, lbl_win_tag, meter_win = _meter_row(ac, '推定勝率(1着になる確率)')
    lbl_top3_val, lbl_top3_tag, meter_top3 = _meter_row(ac, '信頼度＝推定3着内率')

    _stats = tk.Frame(ac, bg=SURFACE)
    _stats.pack(fill='x', padx=px(16), pady=(px(12), 0))
    _stat_lbls = {}
    for _i, _k in enumerate(('勝率', '複勝率', '統計%', '人気', 'オッズ')):
        _cell = tk.Frame(_stats, bg=SURFACE2, padx=px(6), pady=px(4))
        _cell.grid(row=0, column=_i, sticky='nsew', padx=(0 if _i == 0 else px(4), 0))
        _stats.grid_columnconfigure(_i, weight=1)
        tk.Label(_cell, text=_k, bg=SURFACE2, fg=MUTED, font=F['label']).pack()
        _v = tk.Label(_cell, text='–', bg=SURFACE2, fg=TEXT, font=F['bodyb'])
        _v.pack()
        _stat_lbls[_k] = _v
    lbl_comment = tk.Label(ac, text='', bg=SURFACE, fg=TEXT2, font=F['small'], justify='left',
                           anchor='nw', wraplength=px(344))
    lbl_comment.pack(fill='both', expand=True, padx=px(16), pady=(px(10), px(12)))

    bet_card = _card(lower)
    bet_card.pack(side='left', fill='both', expand=True)
    bc = bet_card.inner
    bet_nb = ttk.Notebook(bc, style='Bado.TNotebook')
    bet_nb.pack(fill='both', expand=True, padx=px(2), pady=(px(2), px(4)))

    # 推奨買い目(まとめ)タブ
    _sum_f = tk.Frame(bet_nb, bg=SURFACE)
    bet_nb.add(_sum_f, text='  推奨買い目  ')
    tv_sum = ttk.Treeview(_sum_f, style='Horse.Treeview', show='tree headings',
                          columns=('buy', 'odds', 'hit', 'ev', 'note'))
    tv_sum.heading('#0', text='券種', anchor='w')
    tv_sum.column('#0', width=_colw(170, '券種', '▾ 勝負レース ◎→○▲△'), anchor='w', stretch=False)
    for _c, _h, _w, _a, _smp in (('buy', '買い目', 70, 'center', '12-10'), ('odds', 'オッズ', 70, 'e', '相手99.9倍'),
                                 ('hit', '的中率', 66, 'e', '複勝62%'), ('ev', '期待値', 58, 'e', '188'),
                                 ('note', '条件・備考', 230, 'w', '')):
        tv_sum.heading(_c, text=_h, anchor=_a)
        _cw = _colw(_w, _h, _smp)
        tv_sum.column(_c, width=_cw, minwidth=_cw if _c != 'note' else px(120), anchor=_a, stretch=(_c == 'note'))
    _sum_sb = ttk.Scrollbar(_sum_f, orient='vertical', command=tv_sum.yview, style='Bado.Vertical.TScrollbar')
    tv_sum.configure(yscrollcommand=_sum_sb.set)
    _sum_sb.pack(side='right', fill='y', pady=px(6))
    tv_sum.pack(fill='both', expand=True, padx=(px(6), 0), pady=px(6))
    tv_sum.tag_configure('section', background=P['SECTION_BG'], foreground=BRAND if GUI_THEME != 'dark' else GOLD,
                         font=F['bodyb'])
    tv_sum.tag_configure('gz', foreground=P['GZ'], font=F['bodyb'])
    tv_sum.tag_configure('gzbg', background=P['GZ_BG'])
    tv_sum.tag_configure('stable', foreground=P['GOOD'], font=F['bodyb'])
    tv_sum.tag_configure('muted', foreground=MUTED)
    tv_sum.tag_configure('normal', foreground=TEXT)

    def _make_bet_tab(parent, title):
        f = tk.Frame(parent, bg=SURFACE)
        parent.add(f, text=f'  {title}  ')
        tv = ttk.Treeview(f, style='Horse.Treeview', show='headings',
                          columns=('pair', 'hitrate', 'odds', 'ev', 'note'))
        for _c, _h, _w, _a, _smp in (('pair', '馬番/組', 110, 'center', '12→10'), ('hitrate', '的中率%', 80, 'e', '62.8'),
                                     ('odds', 'オッズ', 80, 'e', '155.0'), ('ev', '期待値', 70, 'e', '1234'),
                                     ('note', '備考', 220, 'w', '')):
            tv.heading(_c, text=_h, anchor=_a)
            _cw = _colw(_w, _h, _smp)
            tv.column(_c, width=_cw, minwidth=_cw if _c != 'note' else px(120), anchor=_a, stretch=(_c == 'note'))
        tv.tag_configure('ev_high', foreground=P['WARN'], font=F['bodyb'])
        tv.tag_configure('hole_bet', foreground=P['ANA'])
        tv.tag_configure('normal', foreground=TEXT)
        tv.tag_configure('gz_bet', foreground=P['GZ'], background=P['GZ_BG'], font=F['bodyb'])
        tv.pack(fill='both', expand=True, padx=px(6), pady=px(6))
        return tv

    tv_tan  = _make_bet_tab(bet_nb, '単勝')
    tv_fuku = _make_bet_tab(bet_nb, '複勝')
    tv_uren = _make_bet_tab(bet_nb, '馬連')
    tv_wide = _make_bet_tab(bet_nb, 'ワイド')
    tv_utan = _make_bet_tab(bet_nb, '馬単')
    tv_hole = None   # 穴馬複勝は生成ロジック側で常に空のためタブを廃止(v136_016)

    # ─── 内部ロジック ─────────────────────────────────────────
    def _set_status(msg):
        var_status.set(msg)

    def _open_deba():
        d = filedialog.askdirectory(title='出馬表HTML(R01_競馬場_日付.html)の入ったフォルダを選択')
        if not d:
            return
        try:
            _races = stat_load_deba_folder(d)
        except Exception as _e:
            messagebox.showerror('出馬表HTMLの読込エラー', str(_e))
            return
        lbl_deba.config(text=_fit_text('統計予想: ' + stat_summary(), F['small'], px(304)))
        _msg = f'出馬表HTML {len(_races)} レースを読み込みました'
        if not stat_model():
            _msg += '(bado_stat_model.json が無いため統計予想は計算できません)'
        _set_status(_msg)
        if var_csv_path.get():
            _load_csv(var_csv_path.get())    # 読込済みのCSVを統計予想つきで再分析

    def _open_csv():
        path = filedialog.askopenfilename(
            title='単勝オッズCSV(nar_shutuba_all_YYYYMMDD.csv)を選択',
            filetypes=[('CSV files', '*.csv'), ('All files', '*.*')])
        if path:
            var_csv_path.set(path)
            _set_status(f'{Path(path).name} を読み込み中…')
            app.update_idletasks()
            _load_csv(path)

    def _choose_outdir():
        d = filedialog.askdirectory(title='出力先フォルダを選択')
        if d:
            var_out_dir.set(d)

    def _race_parts(name):
        venue = str(name[0])
        rno = str(name[2]).translate(_Z2H_DIGITS)
        rno = rno if rno.endswith('R') else f'{rno}R'
        tm = str(name[3])
        tm_short = tm[:5] if len(tm) >= 5 and tm[2:3] == ':' else tm
        return venue, rno, str(name[1]), tm_short

    def _grade_of(analysis):
        _g = str(analysis.get('judgment_class', '')).split(' / ')[0]
        return (_g.replace('級軸', '').strip() or 'D')[:1].upper()

    def _race_cat_of(analysis):
        _jc = str(analysis.get('judgment_class', ''))
        return _jc.split(' / ')[-1].strip() if ' / ' in _jc else ''

    def _has_star(analysis):
        return bool(analysis.get('shobu'))

    def _update_race_row(iid):
        name = _race_name_map.get(iid)
        a = _cache.get(name)
        if name is None or not race_tree.exists(iid):
            return
        venue, rno, dist, tm = _race_parts(name)
        idx = _race_order.index(iid) if iid in _race_order else 0
        zebra = 'even' if idx % 2 == 0 else 'odd'
        if a is None:
            race_tree.item(iid, values=(f'{venue} {rno}', tm, '…', '', ''), tags=(zebra, 'pending'))
            return
        try:
            _sc = calculate_confidence_score(a)['score']
        except Exception:
            _sc = ''
        star = _has_star(a)
        _race_star[iid] = star
        race_tree.item(iid, values=(f'{venue} {rno}', tm, _grade_of(a), _sc, '★' if star else ''),
                       tags=(zebra,) + (('gz',) if star else ()))

    def _apply_filter():
        vsel = var_venue.get()
        star_only = var_star_only.get()
        shown = 0
        for i, iid in enumerate(_race_order):
            name = _race_name_map[iid]
            ok = (vsel in ('', 'すべて') or str(name[0]) == vsel) and (not star_only or _race_star.get(iid))
            if ok:
                race_tree.reattach(iid, '', shown)
                shown += 1
            else:
                race_tree.detach(iid)
        total = len(_race_order)
        var_race_count.set(f'{shown} / {total} R' if shown != total else f'{total} R')

    def _precompute(token, i=0):
        """読込後、全レースを1件ずつ裏で分析して一覧に級・信頼度・★を出す(画面は止めない)。"""
        if token != _load_token[0] or _grouped_data[0] is None:
            return
        if i >= len(_race_order):
            var_progress.set(f'分析完了 {len(_race_order)} R')
            # ★v136_021: データ確認(horselist結合率・騎手変更・取消・分析失敗)
            try:
                _q = data_quality_summary(list(_cache.values()),
                                          n_failed=sum(1 for _v in _cache.values() if _v is None))
                _ql = data_quality_lines(_q)
                if _ql:
                    (messagebox.showwarning if _q['severe'] else messagebox.showinfo)(
                        'データ確認', '\n'.join(_ql[:18]))
            except Exception as _e_dq:
                print(f'  [データ確認] 失敗: {_e_dq}')
            if var_star_only.get():
                _apply_filter()
            return
        iid = _race_order[i]
        name = _race_name_map[iid]
        if name not in _cache:
            try:
                _cache[name] = generate_analysis(_grouped_data[0].get_group(name).copy())
            except Exception as _e:
                print(f'  [GUI] 事前分析に失敗: {name}: {_e}')
                _cache[name] = None
        if _cache.get(name) is not None:
            _update_race_row(iid)
        var_progress.set(f'分析中 {i + 1} / {len(_race_order)}')
        app.after(1, lambda: _precompute(token, i + 1))

    def _load_csv(path):
        _cache.clear()
        stat_set_target_date_from_name(Path(path).name)   # ★v136_019
        _current_analysis[0] = None
        _current_race_name[0] = None
        _load_token[0] += 1
        race_tree.delete(*race_tree.get_children())
        _race_name_map.clear()
        _race_order.clear()
        _race_star.clear()
        _clear_detail()
        p = Path(path)
        lbl_csv_name.config(text=_fit_text(p.name, F['bodyb'], px(304)))
        lbl_csv_dir.config(text=_fit_text(str(p.parent), F['small'], px(304)))
        try:
            df = read_racecard_csv(path)
            lbl_deba.config(text=_fit_text('統計予想: ' + stat_summary(), F['small'], px(304)))   # 自動読込の反映
            df = _sort_by_race_no(df)
            grouped = df.groupby(_RACE_GROUP_KEYS, sort=False)
            _grouped_data[0] = grouped
            venues = []
            for i, name in enumerate(grouped.groups.keys()):
                iid = f'i{i}'
                _race_name_map[iid] = name
                _race_order.append(iid)
                race_tree.insert('', 'end', iid=iid, values=('', '', '', '', ''))
                _update_race_row(iid)
                if str(name[0]) not in venues:
                    venues.append(str(name[0]))
            venue_cb.config(values=['すべて'] + venues)
            var_venue.set('すべて')
            _apply_filter()
            _set_status(f'{p.name} ─ {grouped.ngroups} レースを読み込みました。一覧からレースを選んでください。')
            tok = _load_token[0]
            app.after(50, lambda: _precompute(tok, 0))
            if _race_order:
                race_tree.selection_set(_race_order[0])
                race_tree.focus(_race_order[0])
        except Exception as e:
            messagebox.showerror('読み込みエラー', str(e))
            _set_status(f'読み込みエラー: {e}')

    def _on_race_select(event=None):
        sel = race_tree.selection()
        if not sel:
            return
        name = _race_name_map.get(sel[0])
        if name is None:
            return
        _current_race_name[0] = name
        if _cache.get(name) is None:
            if _grouped_data[0] is None:
                return
            try:
                _cache[name] = generate_analysis(_grouped_data[0].get_group(name).copy())
            except Exception as e:
                messagebox.showerror('分析エラー', str(e))
                return
            _update_race_row(sel[0])
        _current_analysis[0] = _cache[name]
        _refresh_header()
        _refresh_table()
        _refresh_axis_panel()
        _refresh_bet_tabs()

    race_tree.bind('<<TreeviewSelect>>', _on_race_select)

    def _clear_detail():
        lbl_race_title.config(text='レースを選択してください')
        lbl_race_meta.config(text='')
        for w in chip_row.winfo_children():
            w.destroy()
        tree_table.delete(*tree_table.get_children())
        _refresh_axis_panel()
        _refresh_bet_tabs()

    def _refresh_header():
        a = _current_analysis[0]
        name = _current_race_name[0]
        for w in chip_row.winfo_children():
            w.destroy()
        if a is None or name is None:
            return
        venue, rno, dist, tm = _race_parts(name)
        n_heads = len(a['table']) if a.get('table') is not None else 0
        lbl_race_title.config(text=f'{venue}  {rno}')
        lbl_race_meta.config(text=f'{_dist_txt(dist)} ・ 発走 {tm or "-"} ・ {n_heads}頭')
        g = _grade_of(a)
        gbg, gfg = _GUI_GRADE_STYLE.get(g, _GUI_GRADE_STYLE['D'])
        _gdesc = str(a.get('axis_class', '') or f'{g}級軸')
        _chip(chip_row, _gdesc, gbg, gfg).pack(side='left')
        try:
            rec = calculate_confidence_score(a)
            _chip(chip_row, f'信頼度 {rec["score"]} ・ {rec["category"]}',
                  _GUI_TIER_COLOR.get(rec['category'], MUTED)).pack(side='left', padx=(px(8), 0))
        except Exception:
            pass
        cat = _race_cat_of(a)
        if cat:
            _chip(chip_row, cat, _GUI_CAT_COLOR.get(cat, MUTED)).pack(side='left', padx=(px(8), 0))
        _arare = a.get('arare_class', '')
        if _arare:
            _chip(chip_row, f'荒れ度 {a.get("payout_level_score", "")} ・ {_arare}', SURFACE2, TEXT2,
                  font=F['small']).pack(side='left', padx=(px(8), 0))
        if a.get('shobu'):
            _chip(chip_row, '★ %s(単指数%.1f)' % (SHOBU_LABEL, float(a.get('shobu_idx') or 0)), P['GZ']).pack(side='left', padx=(px(8), 0))
        elif a.get('force_reference'):
            _chip(chip_row, '参考のみ(実弾なし)', SURFACE2, MUTED, font=F['small']).pack(side='left', padx=(px(8), 0))

    def _fmt(v, nd=1, dash='-'):
        try:
            fv = float(v)
            if fv != fv:
                return dash
            return f'{fv:.{nd}f}' if nd else f'{int(round(fv))}'
        except Exception:
            return dash

    def _on_heading(col):
        if _sort_state['col'] == col:
            _sort_state['asc'] = not _sort_state['asc']
        else:
            _sort_state['col'] = col
            # 人気・オッズ・馬番は小さい順、それ以外は大きい順が既定
            _sort_state['asc'] = col in ('pop', 'odds', 'ban')
        _refresh_table()

    def _refresh_table():
        for _cid, _hd, _w, _an, _key in TV:
            arrow = ''
            if _cid == _sort_state['col']:
                arrow = ' ▲' if _sort_state['asc'] else ' ▼'
            tree_table.heading(_cid, text=_hd + arrow)
        tree_table.delete(*tree_table.get_children())
        a = _current_analysis[0]
        if a is None:
            return
        df = a['table'].copy()
        if df.empty:
            return
        _key = next((k for c, _h, _w, _an, k in TV if c == _sort_state['col']), '単勝確率')
        if _key in df.columns:
            _s = pd.to_numeric(df[_key], errors='coerce')
            if _key in ('単勝オッズ', '単勝人気'):
                _s = _s.where(_s > 0)          # 欠損(0)は最後に
            df = df.assign(_sk=_s).sort_values('_sk', ascending=_sort_state['asc'],
                                                na_position='last', kind='mergesort')
        for i, (_, row) in enumerate(df.iterrows()):
            mark = str(row.get('印', '') or '')
            base_mark = mark.replace('穴', '')
            odds_v = pd.to_numeric(row.get('単勝オッズ'), errors='coerce')
            vals = (
                mark,
                _fmt(row.get('番'), 0, ''),
                str(row.get('馬名', '')),
                str(row.get('騎手', '')),
                _fmt(row.get(KISHU_DEV_COL), 1, ''),
                str(row.get('展開', '')),
                _fmt(row.get('単勝人気'), 0),
                (_fmt(odds_v, 1) if (odds_v == odds_v and odds_v > 0) else '-'),
                _fmt(row.get('単勝確率'), 1),
                _fmt(row.get('複勝確率'), 1),
                _fmt(row.get('単勝期待値'), 0),
                _fmt(row.get('前売オッズ'), 1),
                _fmt(row.get('人気オッズ'), 1),
                _fmt(row.get('単指数'), 1),
                _fmt(row.get('紐馬指数'), 1),
                str(row.get('統計印', '') or ''),
                _fmt(row.get('統計勝率'), 1, ''),
                _fmt(row.get('オッズ勝率'), 1, ''),
            )
            if base_mark.startswith('◎'):
                tags = ['axis']
            elif base_mark[:1] in ('○', '▲', '△', '☆'):
                tags = ['mark']
            else:
                tags = ['even' if i % 2 == 0 else 'odd']
            if '穴' in mark:
                tags.append('ana')
            tree_table.insert('', 'end', values=vals, tags=tuple(tags))

    def _refresh_axis_panel():
        a = _current_analysis[0]
        if a is None:
            grade_badge.config(text='–', bg=BORDER, fg='#FFFFFF')
            lbl_axis_name.config(text='─')
            lbl_grade_desc.config(text='')
            for _l in (lbl_win_val, lbl_top3_val):
                _l.config(text='–')
            lbl_win_tag.config(text='')
            lbl_top3_tag.config(text='')
            meter_win.set(0)
            meter_top3.set(0)
            for _v in _stat_lbls.values():
                _v.config(text='–')
            lbl_comment.config(text='')
            return
        tbl = a['table']
        axis_row = tbl[tbl['印'].astype(str).str.startswith('◎')]
        g = _grade_of(a)
        gbg, gfg = _GUI_GRADE_STYLE.get(g, _GUI_GRADE_STYLE['D'])
        grade_badge.config(text=g, bg=gbg, fg=gfg)
        if not axis_row.empty:
            r0 = axis_row.iloc[0]
            lbl_axis_name.config(text=f'{_fmt(r0.get("番"), 0, "")}  {r0.get("馬名", "")}')
            _stat_lbls['勝率'].config(text=_fmt(r0.get('単勝確率'), 1) + '%')
            _stat_lbls['複勝率'].config(text=_fmt(r0.get('複勝確率'), 1) + '%')
            _stat_lbls['統計%'].config(text=_fmt(r0.get('統計勝率'), 1))
            _stat_lbls['人気'].config(text=_fmt(r0.get('単勝人気'), 0))
            _ov = pd.to_numeric(r0.get('単勝オッズ'), errors='coerce')
            _stat_lbls['オッズ'].config(text=(_fmt(_ov, 1) + '倍') if (_ov == _ov and _ov > 0) else '-')
        else:
            lbl_axis_name.config(text='─')
        lbl_grade_desc.config(text=str(a.get('axis_class', '')))
        rec = calculate_confidence_score(a)
        _tc = _GUI_TIER_COLOR.get(rec['category'], BRAND2)
        _ew = a.get('judge_est_win')
        if _ew is not None:
            lbl_win_val.config(text=f'{_ew:.0f}%')
            meter_win.set(_ew, gbg if g != 'D' else MUTED)
            lbl_win_tag.config(text=f'{g}級', fg=gbg if g not in ('S',) else GOLD2)
        else:
            lbl_win_val.config(text=_fmt(a.get('axis_win_prob'), 0) + '%')
            meter_win.set(a.get('axis_win_prob', 0), BRAND2)
            lbl_win_tag.config(text='モデル')
        lbl_top3_val.config(text=f'{rec["score"]}%' if a.get('judge_est_top3') is not None else f'{rec["score"]}')
        lbl_top3_tag.config(text=rec['category'], fg=_tc)
        meter_top3.set(rec['score'], _tc)
        lbl_comment.config(text=str(a.get('judgment_comment', '') or ''))

    def _fill_bet_tv(tv, recs, kind='single', yuryoku=None, myomi=None,
                     force_ref=False, kind_is_umatan=False, kind_s=None):
        if tv is None:
            return
        if kind_s is None:
            kind_s = 'umatan' if kind_is_umatan else 'umaren'
        tv.delete(*tv.get_children())
        hole_bans = set(_current_analysis[0].get('hole_bans', [])) if _current_analysis[0] else set()
        for i, rec in enumerate(recs):
            _cn = ''
            if kind == 'single':
                pair = str(rec.ban)
                note = '参考' if force_ref else ''
                is_hole = rec.ban in hole_bans
            elif kind == 'holefuku':
                pair = f"{rec.ban}番 {rec.horse_name}"
                note = '★穴馬軸'
                is_hole = True
            else:
                pair = f"{rec.ban1}{'→' if kind_s == 'umatan' else '-'}{rec.ban2}"
                _k = (rec.ban1, rec.ban2)
                _an = _current_analysis[0] or {}
                _cmap = _an.get({'umatan': 'gz_umatan_cond', 'umaren': 'gz_umaren_cond',
                                 'wide': 'gz_wide_cond'}[kind_s], {})
                _cn = _cmap.get(_k, '')
                _clab = _cn if _cn in GZ_COND_DEFS else ''
                _st = _an.get('gz_stakes', {}).get((rec.ban1, rec.ban2, kind_s), 0)
                if force_ref:
                    note = ('%s 参考(見送り)' % _clab).strip()
                elif _cn and _st > 0:
                    note = '%s ★採用 %d円' % (_clab, _st)
                elif _cn:
                    note = '%s 見送り(点数上限)' % _clab
                else:
                    note = '参考'
                is_hole = False
            ev = rec.ev
            tag = 'gz_bet' if _cn else ('hole_bet' if is_hole else ('ev_high' if ev >= 110 else 'normal'))
            tv.insert('', 'end', values=(pair, _fmt(rec.hit_rate, 1), _fmt(rec.odds, 1),
                                         _fmt(ev, 0), note), tags=(tag,))

    def _fill_summary(a):
        tv_sum.delete(*tv_sum.get_children())
        if a is None:
            return
        force_ref = bool(a.get('force_reference', False))
        stakes = a.get('gz_stakes', {}) or {}
        sep_of = {'umatan': '→', 'umaren': '-', 'wide': '-'}
        jp_of = {'umatan': '馬単', 'umaren': '馬連', 'wide': 'ワイド'}
        any_row = False

        # 1) 本線枠(実弾)
        main_rows = []
        for kind, key_recs, key_cond in (('wide', 'wide_recs', 'gz_wide_cond'),
                                         ('umaren', 'umaren_recs', 'gz_umaren_cond'),
                                         ('umatan', 'umatan_recs', 'gz_umatan_cond')):
            cmap = a.get(key_cond, {}) or {}
            for rec in a.get(key_recs, []) or []:
                cn = cmap.get((rec.ban1, rec.ban2), '')
                if not cn:
                    continue
                st = stakes.get((rec.ban1, rec.ban2, kind), 0)
                lab = _NAR_GZ_LABEL.get(cn, (cn, '', '', ''))
                code = str(lab[0]).split(' ')[0]
                if gz_tier(cn) != 'main':
                    code += '(%s)' % GZ_TIER_LABEL.get(gz_tier(cn), '')
                if _stat_ura_of(a, cn):   # ★v136_020
                    code += ' 【' + STAT_URA_LABEL + '】'
                if force_ref:
                    note, tg = f'{code}  参考(見送り)', ('muted',)
                elif st > 0:
                    note, tg = f'{code}  ★採用 {st}円', ('gz', 'gzbg')
                else:
                    note, tg = f'{code}  見送り(点数上限)', ('muted',)
                main_rows.append((jp_of[kind], f'{rec.ban1}{sep_of[kind]}{rec.ban2}',
                                  _fmt(rec.odds, 1) + '倍', _fmt(rec.hit_rate, 1) + '%', _fmt(rec.ev, 0), note, tg))
        if main_rows:
            sec = tv_sum.insert('', 'end', text='本線枠（実弾）', open=True, tags=('section',),
                                values=('', '', '', '', '参考のみ' if force_ref else ''))
            for r in main_rows:
                tv_sum.insert(sec, 'end', text=r[0], values=r[1:6], tags=r[6])
            any_row = True

        # 2) 安定(単勝・複勝)
        st_rows = []
        for key_rec, key_tag, jp in (('stable_tan_rec', 'stable_tan_tag', '単勝'),
                                     ('stable_fuku_rec', 'stable_fuku_tag', '複勝')):
            rec = a.get(key_rec)
            if rec is None:
                continue
            tagtxt = str(a.get(key_tag, '') or '')
            ok = tagtxt.startswith('安定')
            st_rows.append((jp, str(rec.ban), (_fmt(rec.odds, 1) + '倍') if rec.odds else '-',
                            _fmt(rec.hit_rate, 0) + '%', _fmt(rec.ev, 0) if rec.ev else '-', tagtxt,
                            ('stable',) if ok else ('muted',)))
        if st_rows:
            sec = tv_sum.insert('', 'end', text='単勝・複勝（安定）', open=True, tags=('section',),
                                values=('', '', '', '', ''))
            for r in st_rows:
                tv_sum.insert(sec, 'end', text=r[0], values=r[1:6], tags=r[6])
            any_row = True

        # 3) 次点本線枠(表示のみ)
        jrecs = a.get('jiten_recs') or []
        if jrecs:
            sec = tv_sum.insert('', 'end', text='次点本線枠（表示のみ）', open=True, tags=('section',),
                                values=('', '', '', '', ''))
            for jr in jrecs:
                _sep = sep_of.get(jr.get('kind'), '-')
                for p in jr.get('partners', []):
                    _od = p.get('odds')
                    tv_sum.insert(sec, 'end', text=jp_of.get(jr.get('kind'), jr.get('kind')),
                                  values=(f"{jr['axis_ban']}{_sep}{p['ban']}",
                                          (f'相手{_od:.1f}倍' if _od else '-'),
                                          f"複勝{p.get('place', 0):.0f}%", '-',
                                          f"{jr.get('tag', '')} 検証ROI{jr.get('roi', 0):.0f}%・的中{jr.get('hit', 0):.0f}%・N{jr.get('n', 0)}"
                                          + (' 【' + STAT_URA_LABEL + '】' if _stat_ura_of(a, jr.get('tag')) else '')),
                                  tags=('normal',))
            any_row = True

        # ★v136_019: 統計予想(参考・実弾外)
        _stc = a.get('stat')
        if _stc:
            _mk_line = ' '.join('%s%d' % (_stc['marks'][_b], _b) for _b in _stc['order'][:4])
            sec = tv_sum.insert('', 'end', text='統計予想（参考・実弾外）', open=True, tags=('section',),
                                values=('', '', '', '', _mk_line))
            for _x in _stc['bets']:
                tv_sum.insert(sec, 'end', text=jp_of.get(_x['kind'], _x['kind']),
                              values=(f"{_x['ban1']}-{_x['ban2']}", '-', f"{_x['prob']:.1f}%", '-',
                                      f"統計%: {_stc['p_stat'][_x['ban1']]:.1f} / {_stc['p_stat'][_x['ban2']]:.1f}"),
                              tags=('normal',))
            any_row = True
        # ★v136_020: 統計裏付け・観察(統計紐・記録のみ)
        _ura = a.get('stat_ura') or {}
        _obs_ok = [o for o in (a.get('stat_obs') or []) if o['ok']]
        if _ura or _obs_ok:
            sec = tv_sum.insert('', 'end', text='統計裏付け・観察（記録のみ）', open=True, tags=('section',),
                                values=('', '', '', '', '軸の統計勝率50%以上=統計裏付け'))
            _seen = {}
            for _k in STAT_AXIS_CODES:
                if _k in _ura:
                    _seen.setdefault(_ura[_k]['ban'], []).append(_k)
            for _b, _ks in _seen.items():
                _u = _ura[_ks[0]]
                tv_sum.insert(sec, 'end', text='軸',
                              values=(f'{_b}番', '-', f"統計{_u['ps']:.1f}%", '-',
                                      '・'.join(_ks) + ('  ★' + STAT_URA_LABEL if _u['ok'] else '  裏付けなし')),
                              tags=('gz',) if _u['ok'] else ('muted',))
            for o in _obs_ok:
                tv_sum.insert(sec, 'end', text=jp_of.get(o['kind'], o['kind']),
                              values=(f"{o['axis_ban']}-{o['ban']}",
                                      (f"相手{o['odds']:.1f}倍" if o.get('odds') else '-'),
                                      f"紐統計{o['p_pt']:.1f}%", '-',
                                      f"{o['tag']} {o['label']}（{o['n']}点 的中{o['hit']:.0f}% 回収{o['roi']:.0f}%）"),
                              tags=('normal',))
            any_row = True

        # 4) 参考予想 ◎→○▲△(表示のみ)
        refs = a.get('reference_bets') or []
        if refs:
            tbl = a.get('table')
            _mk_of = {}
            _nm_of = {}
            try:
                for _, _r in tbl.iterrows():
                    _mk_of[int(_r['番'])] = str(_r.get('印', '') or '').replace('穴', '')
                    _nm_of[int(_r['番'])] = str(_r.get('馬名', '') or '')
            except Exception:
                pass
            _rdesc = str(refs[0].get('desc', '') or '')
            _rdesc = _rdesc.replace('◎→○▲△', '').strip('()（） ')
            sec = tv_sum.insert('', 'end', text=('%s ◎→○▲△' % SHOBU_LABEL) if a.get('shobu') else '参考予想 ◎→○▲△', open=True, tags=('section',),
                                values=('', '', '', '', _rdesc))
            for rf in refs:
                _sep = sep_of.get(rf.get('kind'), '-')
                for b in rf.get('partner_bans', []):
                    _m = _mk_of.get(int(b), '')
                    tv_sum.insert(sec, 'end', text=jp_of.get(rf.get('kind'), rf.get('kind')),
                                  values=(f"{rf['axis_ban']}{_sep}{b}", '', '', '',
                                          f"{_m or '・'} {b} {_nm_of.get(int(b), '')}"),
                                  tags=('normal',))
            any_row = True

        if not any_row:
            tv_sum.insert('', 'end', text='買い目なし', values=('', '', '', '', 'このレースは買い目なし'),
                          tags=('muted',))

    def _refresh_bet_tabs():
        a = _current_analysis[0]
        if a is None:
            for tv in (tv_sum, tv_tan, tv_fuku, tv_uren, tv_wide, tv_utan):
                tv.delete(*tv.get_children())
            return
        _fr = a.get('force_reference', False)
        _fill_summary(a)
        _fill_bet_tv(tv_tan,  a.get('tan_recs', []),  'single', force_ref=_fr)
        _fill_bet_tv(tv_fuku, a.get('fuku_recs', []), 'single', force_ref=_fr)
        # ★v134: 安定運用の単勝/複勝を各タブへ追記(備考=安定タグ)
        for tv, key_rec, key_tag, dflt in ((tv_tan, 'stable_tan_rec', 'stable_tan_tag', '安定(単)'),
                                          (tv_fuku, 'stable_fuku_rec', 'stable_fuku_tag', '安定(複)')):
            rec = a.get(key_rec)
            if rec is not None:
                _tag = a.get(key_tag, dflt)
                tv.insert('', 'end', values=(rec.ban, _fmt(rec.hit_rate, 1), _fmt(rec.odds, 1),
                                             _fmt(rec.ev, 0), _tag),
                          tags=('ev_high' if '安定' in _tag else 'normal',))
        _fill_bet_tv(tv_utan, a.get('umatan_recs', []), 'pair', force_ref=_fr, kind_is_umatan=True, kind_s='umatan')
        _fill_bet_tv(tv_uren, a.get('umaren_recs', []), 'pair', force_ref=_fr, kind_s='umaren')
        _fill_bet_tv(tv_wide, a.get('wide_recs', []),   'pair', force_ref=_fr, kind_s='wide')

    def _run_analysis():
        csv_path = var_csv_path.get()
        if not csv_path or not Path(csv_path).exists():
            messagebox.showwarning('ファイル未選択', '単勝オッズCSVを選択してください。')
            return
        out_dir = Path(var_out_dir.get())
        out_dir.mkdir(parents=True, exist_ok=True)

        # 入力ファイル名からレース日付を抽出（例: updatedtihou0526.csv → 20260526）
        # 8桁(YYYYMMDD)優先、なければ4桁(MMDD)に現在の年を付与
        import re
        fname = Path(csv_path).stem
        m8 = re.search(r'(20\d{6})', fname)
        if m8:
            today_str = m8.group(1)
        else:
            m4 = re.search(r'(\d{4})', fname)
            if m4:
                mmdd = m4.group(1)
                if len(mmdd) == 4 and 1 <= int(mmdd[:2]) <= 12 and 1 <= int(mmdd[2:]) <= 31:
                    today_str = f"{datetime.now().year}{mmdd}"
                else:
                    today_str = date.today().strftime('%Y%m%d')
            else:
                today_str = date.today().strftime('%Y%m%d')

        output_csv = out_dir / f'race_forecasts_{today_str}.csv'
        excel_path = out_dir / f'race_forecasts_{today_str}_colored.xlsx'

        _set_status('分析中・ファイル出力中…')
        btn_run.set_enabled(False)
        app.update_idletasks()

        def _worker():
            try:
                if _grouped_data[0] is None:
                    app.after(0, lambda: messagebox.showwarning('未読込', 'まずCSVを読み込んでください。'))
                    return
                _run_output(_grouped_data[0], output_csv, excel_path)
                app.after(0, lambda: _set_status(f'出力完了: {output_csv.name} / {excel_path.name}'))
                app.after(0, lambda: messagebox.showinfo('完了', f'ファイル出力が完了しました。\n\n{output_csv}\n{excel_path}'))
            except Exception as e:
                _err = str(e)
                app.after(0, lambda m=_err: messagebox.showerror('出力エラー', m))
                app.after(0, lambda m=_err: _set_status(f'エラー: {m}'))
            finally:
                app.after(0, lambda: btn_run.set_enabled(True))

        threading.Thread(target=_worker, daemon=True).start()

    def _init_sash():
        try:
            h = paned.winfo_height()
            if h > 200:
                paned.sashpos(0, max(px(220), h - px(336)))
        except Exception:
            pass
    app.after(300, _init_sash)

    # コメント欄の折り返し幅をカード幅に合わせる
    ac.bind('<Configure>', lambda e: lbl_comment.config(wraplength=max(px(200), e.width - px(36))))
    # ショートカット
    app.bind_all('<Control-o>', lambda e: _open_csv())
    app.bind_all('<F5>', lambda e: _run_analysis())

    # コマンドライン引数でCSVが指定されていれば自動読込
    if _args and _args.data and Path(str(_args.data)).exists():
        var_csv_path.set(str(_args.data))
        var_out_dir.set(str(_args.output))
        app.after(200, lambda: _load_csv(str(_args.data)))

    app.mainloop()


# ══════════════════════════════════════════════════════════════
# エントリーポイント
# ══════════════════════════════════════════════════════════════
def main():
    global _args
    _args = _parse_args()
    # ★v136_003: 旧・参考枠(gz_reference.py + 有望条件.json)は全廃したため、
    #   ここでの読込処理は不要になった(次点本線枠のJSON読込は別経路で継続)。
    if getattr(_args, 'deba', None):
        try:
            stat_load_deba_folder(_args.deba)
        except Exception as _e:
            print(f'[統計予想] --deba の読込に失敗: {_e}')
    run_gui()


if __name__ == '__main__':
    main()