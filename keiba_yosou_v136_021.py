# =====================================================
# 地方競馬(NAR)予想 v136_021
#
# ★v136_021: 前提データの取り違えを検出・補正する修正(買い目の条件そのものは v136_020 と同じ)。
#   1) horselist の結合確認: 毎回ファイルの更新を見て読み直す(旧版は起動中ずっと最初の内容を使い回していた)。
#      結合率が90%未満なら GUI で警告、予想表HTMLの先頭に赤帯。50%未満のレースにはレース内に注意書き。
#   2) 騎手変更: horselist の騎手名が予想CSVと違う馬(略称の違いは先頭2文字一致で同一騎手とみなす)は、
#      騎手を新しい騎手に置き換え(表示は『野畑凌※』)、騎手指数は同じ日・同じ場のCSVにある新騎手の値(中央値)に差し替える。
#      新騎手がCSVに居なければ騎手指数は旧騎手のまま(注意書きに明記)。
#   3) 出走取消・除外: 出馬表HTMLを読み込んでいれば、『出走取消/競走除外』の馬を予想から除いて確率を計算し直す。
#   4) テキストCSVで WD1 が『馬連』と表示されていたのを修正。ワイドだけのレースでも採用条件の見出しを出す。
#   5) 分析エラーのレースは予想表HTMLの先頭に一覧で表示(旧版は黙って消えていた)。トレースバックは全件出す。
#   6) その他: 次点の検証成績を keiba_tool が読める書式に、Excelの色付けの1列ずれ、買い目なし判定にワイド、
#      単勝確率が欠損した馬があっても融合勝率が壊れないように、一括予想の対象日の決め方をGUIと統一。
#   ・データ確認の結果は 予想表HTML先頭 / レース内の注意書き / Excel・CSV / GUIのダイアログ / 統計観察記録CSV に出る。
#
# ★v136_020: 『統計裏付け』の印と、統計1位を紐にする観察条件(ST1〜ST6)の記録を追加(いずれも実弾外)。
#   出典: 紐条件探索5_統計1位紐_0716-0913.xlsx(7/16〜9/13・統計勝率は月別horselistから復元した値)
#   ・統計裏付け = 本線/次点の軸馬(UR1・UR2・WD1・OU1 の軸)の統計勝率が 50% 以上。
#       馬表の統計印に『裏』を付け、その軸の本線・次点の買い目に【統計裏付け】を表示する。
#       (7/16〜9/13: 軸の2着内率 UR1軸 88.9% vs 73.9%、WD1軸 90.3% vs 78.8%、OU1軸 87.7% vs 77.8%)
#   ・観察条件(表示・記録のみ。紐=軸を除いた統計勝率1位の1頭。条件を満たさなければ見送り・繰り上げなし)
#       ST1 UR1軸 馬連  紐の統計勝率10%以上 & 軸の統計勝率40%以上      35点 的中57.1% 回収145%
#       ST2 UR1軸 ワイド 紐の複勝確率40以上 & 統計勝率≧単勝確率(正規化)  35点 的中82.9% 回収118%
#       ST3 WD1軸 馬連  紐の複勝確率40以上 & 統計勝率10%以上          32点 的中56.2% 回収161%
#       ST4 WD1軸 ワイド 紐の複勝確率45以上 & 統計勝率15%以上          22点 的中95.5% 回収129%
#       ST5 OU1軸 馬連  紐の統計勝率15%以上 & 軸の統計勝率50%以上      30点 的中70.0% 回収142%
#       ST6 OU1軸 ワイド 同上                                         30点 的中90.0% 回収116%
#     ※全期間を見て選んだ値で上振れを含む(前半で選び後半で確かめた実力は回収100〜108%程度)。
#   ・記録: 出力フォルダに <出力名>_統計観察記録.csv(1レース1行: 各軸の統計勝率・裏付け・本線/次点/観察の買い目)。
#   ・本線枠・次点本線枠・予想印・参考予想・統計予想は v136_019 のまま変更なし。
#
# ★v136_019: 出馬表HTML(地方競馬情報サイト R01_高知_20260927.html 等)を読み込み、統計予想を追加。
#   ・GUI『出馬表HTMLフォルダを読込』でフォルダ(サブフォルダ含む)の R01〜R12 を競馬場別に一括読込。
#     一括予想・単発予想のどちらでも、CSVの日付・競馬場・R・馬番(馬名照合つき)で自動的に結合する。
#   ・統計モデル(bado_stat_model.py + bado_stat_model.json、同じフォルダに置く):
#       近5走(地方)の着順率・着差・人気・スピード指数・上がり指数・間隔、着別成績(経験ベイズ縮約)、
#       騎手/調教師の直近365日成績 → 条件付きロジット(Plackett–Luce 上位3着)で『統計勝率』。
#       学習 2025/10〜2026/7/15(月別horselist 着順入り)。
#   ・融合: log融合勝率 = a·log(既存の単勝確率) + b·log(統計勝率) をレース内で正規化(a,b は 7/16〜9/13 で推定)。
#     検証(8/16〜9/13, 1,092R・学習に使っていない期間): 対数損失 既存1.834 → 統計1.784 → 融合1.761、
#       ◎単勝的中 36.3%→37.4%、◎-○▲ 馬連的中 13.8%→15.8%、ワイド的中 30.0%→32.8%(回収率は70〜80%で改善なし)。
#   ・予想表: 馬表に 統計印・統計%・融合% の列、レースごとに『統計予想(参考・実弾外)』
#     (統計印◎○▲△ と ◎-○/◎-▲ の馬連・ワイド、Harville推定の的中確率)を追加。CSV/Excel/GUIにも表示。
#   ・本線枠(UR1/UR2/WD1)・次点本線枠(OW1/OU1)・予想印◎○▲△・参考予想は v136_018 のまま変更なし。
#   ・出馬表HTMLを読み込まない/モデルファイルが無いときは統計列が空になるだけで、従来どおり動く。
#   ・再学習: bado_stat_train.py(月別horselist を追加して実行 → bado_stat_model.json を更新)。
#
# ★v136_018: 本線枠を『BADO 実戦条件シート』の実戦候補3条件へ全面変更。旧本線枠は全廃。
#   出典: 紐条件探索3_単指数90軸_区分別_0716-0923.xlsx / 紐条件探索4_単確50軸_区分別_0716-0923.xlsx
#   ・本線枠(実弾・1点100円)
#       UR1 馬連① 軸=単指数90以上のうち単指数最大の馬、その予想オッズ1.9倍以下
#                 紐=単勝確率8以上のうち紐馬指数1位の1頭           (99R 的中30.3% 回収121.3%)
#       UR2 馬連② 軸=同じ軸馬が単指数90〜93かつ予想オッズ1.4倍以下
#                 紐=単勝人気3以内かつ紐馬指数50以上のうち紐馬指数1位 (39R 的中48.7% 回収156.2%)
#       WD1 ワイド① 軸=補正後単勝確率50%以上のうち最大の馬、その馬が単指数90以上かつ単勝人気1
#                 紐=複勝確率45以上かつ単指数65以上のうち紐馬指数1位 (33R 的中84.8% 回収132.4%)
#     ・UR1とUR2が同じ組になったら1点(UR1扱い)。軸がオッズ等を満たさなければ見送り(繰り上げなし)。
#   ・次点本線枠(表示のみ・実弾外)= シートの『観察のみ』2条件
#       OW1 ワイド 軸=UR1と同じ(単指数90軸・予想オッズ1.9倍以下) 紐=複勝確率40以上&単勝オッズ30倍以内→紐馬指数1位
#       OU1 馬連   軸=補正後単勝確率50%以上の最大・予想オッズ1.9倍以下 紐=単勝確率15以上&紐馬指数60以上→紐馬指数1位
#     ・旧次点(有望条件.json/最適紐008)は使わない。本線と同じ組は次点に出さない。
#   ・旧本線枠(V1〜V7・AR5〜AR7)は廃止(定義はレジストリに残置、GZ_MAIN_KEEP から除外)。
#   ・単勝・複勝【安定】も本線から外した(STABLE_BET_ENABLE=False。True で復活)。
#   ・判定値は予想表に表示される値そのもの(探索 keiba_tool008 と同じ):
#       単勝確率/複勝確率/紐馬指数 = 表示と同じ小数1桁、単指数 = 整数、
#       人気/単オッズ = 騎手補正後(adjusted_popularity / adjusted_odds)。
#     ※旧本線枠は 人気/オッズを補正前の列で判定していたため、探索(表示値)と定義がずれていた。
#   ・予想表 meta の bado-logic に honsen=sheet1 を追加(keiba_tool008 の版区別用)。
#
# ★v136_017: GUI の高DPI表示崩れを修正(予想・買い目・出力は v136_016/015 と同一)。
#   ・原因: フォントはポイント指定で拡大率(125%/150%)に応じて大きくなるのに、列幅・パネル幅・
#     余白はピクセル固定だったため、Windows の拡大表示で列が潰れ・パネルがはみ出していた。
#   ・対策: 拡大率 S(=画面DPI/96) を全ピクセル寸法に掛ける。さらに実質の画面が狭いとき
#     (例 1920x1080 を150%表示)は余白・パネル・文字を一律に詰める(D=0.72〜1.0)+最大化で起動。
#   ・表の列幅は『見出しと代表値が収まる幅』以上を保証(_colw)。ファイル名・出力先は1行で中央省略。
#   ・出力先が相対パス('.')のときも絶対パスで表示。
#
# ★v136_016: GUI のデザインを刷新(予想ロジック・買い目・出力ファイルの中身は v136_015 と同一)。
#   ・配色を研究用HTML/note版と同じ『深緑×クリーム×金』に統一(BADO_GUI_THEME=dark で暗色版)。
#   ・レース一覧に 発走・軸級・信頼度・★採用 を表示し、開催場/★採用のみで絞り込み可能に
#     (読み込み後に裏で全レースを順に分析。画面は操作できる)。
#   ・レース見出しに 軸級/信頼度/レース分類/荒れ度/★採用 をチップ表示。
#   ・出走馬表は研究用HTMLと同じ列構成(軸馬指数を追加、BADO指数/hybrid は非表示)。
#     見出しクリックで並べ替え、◎・○▲△(参考買い目の相手)・穴 を色分け。
#   ・◎軸馬カード: 軸級バッジ + 推定勝率/信頼度(推定3着内率)のメーター + 主要値。
#   ・『推奨買い目』タブに 本線枠(★採用)/安定/次点本線枠/参考予想(◎→○▲△) を一覧表示。
#     券種別タブ(単勝/複勝/馬連/ワイド/馬単)は従来どおり。常に空だった『穴馬複勝』タブは廃止。
#   ・高DPI対応、日本語フォント自動選択、Ctrl+O=CSVを開く / F5=予想実行・ファイル出力。
#   ・古い表記(タイトルの v079 等)を BADO_VERSION に統一。
#
# ★v136_015: 参考予想を回収率重視に変更(REF_VALUE_LAMBDA 1.0 → 4.0)。それ以外は v136_014 と同一。
#   ・○▲△(=参考買い目の相手3頭)の選び方で『配当の妙味』の重みを上げた。構造・◎・点数は不変。
#     本線枠・次点本線枠・安定(単複)・軸級・信頼度・表示の確率/指数は一切変更なし。
#   ・検証(2026/07/16〜09/19 2,493R, ◎から3点, 1点100円):
#       λ=1: 馬連 的中41.6% 回収78.2% / ワイド 的中61.7% 回収76.9% / 合算77.6%
#       λ=4: 馬連 的中37.9% 回収79.8% / ワイド 的中59.3% 回収81.2% / 合算80.5%
#     ワイド回収 +4.3pt[95%CI +0.9,+7.8] は7〜8月(80.6%)・9月(82.5%)・頭数帯別の全てで改善。
#     馬連回収の差(+1.5pt)は誤差の範囲。的中率は 馬連-3.7pt / ワイド-2.4pt。
#     λ=4 は同データのグリッドで選んだ値なので、実際の改善幅はこれより小さい可能性がある。
#   ・λ=1(的中重視)に戻すときは REF_VALUE_LAMBDA = 1.0。
#
# ★v136_014: 予想表に『作成した版』を記録 + note版の表示修正(予想・買い目は v136_013 と同一)。
#   ・研究用HTML / note用HTML の <head> に次の2つの meta を埋め込む(画面には出ない)。
#       <meta name="bado-version" content="v136_014">
#       <meta name="bado-logic" content="judge=new;ref=new">   (USE_NEW_JUDGE / REF_NEW_LOGIC)
#     keiba_tool008 はこれを読み、版・判定ロジックごとに集計/絞り込みできる。
#     meta が無いファイル(v136_013 以前)は『不明(記録なし)』として扱われる。
#   ・note版: 印が『○穴』『▲穴』等のとき ○▲△ が消えて『穴』だけになっていた不具合を修正
#     (印の丸に加えて小さな『穴』バッジを併記)。
#   ・note版: 信頼度・軸級の説明を v136_008 以降の定義に合わせた
#     (信頼度=◎が3着以内に入る推定確率%、S〜D=◎の推定勝率の区分)。
#     REF_NEW_LOGIC / USE_NEW_JUDGE を 0 にしたときは旧説明を出す。
#   ・note版のフッターに版を小さく表示。
#
# ★v136_010: 予想印(○▲△)と参考買い目を再設計。本線枠・次点本線枠・安定(単複)は完全に不変。
#   ・参考買い目 = ◎から馬連・ワイド 各3点(◎-○/▲/△)。予想印と完全連動、☆は廃止。
#   ・○▲△ = ◎との馬連確率(非表示のブレンド勝率から Harville)× 妙味(オッズ比)の上位3頭
#     (ref_partner_order / REF_VALUE_LAMBDA=1.0 ※v136_015で4.0)。◎・確率・指数の表示値は変更なし。
#   ・7/16〜9/19 2,493R(3点固定)で 馬連 的中35.3→41.6% 回収74.6→78.2%、
#     ワイド 的中57.0→61.7% 回収76.5→76.9%。7〜8月/9月のどちらでも的中率は改善。
#   ・既定では本線枠と同じ組も参考に残す(REF_EXCLUDE_MAIN=False。常に3点)。
#   ・不具合修正: 旧ロジックのフォールバック時(該当馬0頭)に参考買い目が全頭流しに
#     なっていた問題を上位3頭に修正(REF_NEW_LOGIC=0 で旧ロジックを使う場合に有効)。
#
# ★v136_009: 新判定の係数を 7/16〜9/19 の 2,324R で再当てはめ + レース分類の2頭軸を停止。
#   ・JUDGE_BLEND_A/B = 0.4317/0.7677, JUDGE_TOP3_SLOPE/INTERCEPT = 0.6374/-0.1356
#     (horselist結合失敗でモデル確率が崩れた 7/22・8/6・8/9・8/25 は当てはめから除外)
#   ・JUDGE_CAT_TWO_ENABLE = False: 9月の未使用データで ◎-対抗 馬連的中 1/20R と機能せず。
#     該当レースは標準レース(または見送り=表示は標準)になる。
#   ・区分境界・表示形式・買い目は v136_008 と同一。
#
# ★v136_008: 軸判定(S〜D級)・信頼度・レース分類のロジックを再設計(表示形式・買い目は不変)。
#   検証(2026/07/16〜08/31 NAR 1,714R)で旧判定は B/C/D級・信頼度 中程度以下がほぼ無差別、
#   レース分類は ◎ と別の馬(複合スコア1位)を旧スケールの式で判定していた。
#   新判定はすべて ◎(アンサンブル単勝確率1位=表示の軸) の推定確率に基づく:
#     ・ブレンド勝率 = 単勝確率^A × 補正後オッズ由来確率^B をレース内正規化(対数線形)
#     ・推定3着内率 = ブレンド勝率から Harville で3着内確率を出し、ロジスティック較正
#     ・軸級   = ◎の推定勝率で S/A/B/C/D
#     ・信頼度 = ◎の推定3着内率(%)をそのまま 0-100 の数値に。区分境界は新設定
#     ・レース分類 = ◎の推定勝率/3着内率 + ◎と対抗の馬連確率(2頭軸)
#   前半で当てはめ後半で検証しても単調・較正(S 複勝圏89.6% … D 49.2%)。
#   旧判定に戻すときは USE_NEW_JUDGE = False(または環境変数 USE_NEW_JUDGE=0)。
#   ※ 信頼度による参考強制(NSL_TIER_REFERENCE_ONLY)の判定は旧式のまま(買い目不変)。
#
# ★v136_007: note 読者向けの予想表を追加(予想ロジック・買い目は v136_006 と同一)。
#   出力: <名前>_note.html / <名前>_note.pdf (A4縦・スマホ幅1列)
#     ・表紙: 日付/開催場/予想レース数/推奨買い目のあるレース数/点数
#     ・本日の推奨買い目一覧(発走順・レースへのリンク付き)
#     ・予想表の見方(用語の説明)
#     ・各レースカード: 研究用HTMLと同じ全項目を4グループで表示
#         馬(印・馬番/馬名・性齢・斤量・騎手/脚質) / オッズ(人気・単勝) /
#         AI予測(勝率・複勝率バー・期待値) / 指数(単指数・複指数・軸馬・紐馬・騎手補正)
#       スマホ幅は印・馬名を固定した横スクロール、PDF(A4縦)は全列が1枚に収まる。
#       見出しにレース分類・軸級・信頼度区分。以下 推奨買い目 → 単勝・複勝 → 参考買い目
#     ・金額・内部用語(本線/stable等)は出さない。末尾に注意書き
#   従来の研究用HTML/PDF/Excel/CSV は一切変更なし(keiba_tool の解析対象のため)。
#   設定: NOTE_HTML_ENABLE / NOTE_SHOW_REFERENCE / NOTE_SHOW_STABLE_OUT /
#         NOTE_SHOW_COND_CODE / NOTE_BRAND / NOTE_DISCLAIMER
#   html_to_pdf に landscape 引数を追加(既定 True=従来どおり横向き)。
#
# ★v136_006: 本線枠を『紐条件アーカイブ第2版』の再判定に合わせて組み替え。
#   再検証の要点: 最適紐016〜026は全て同一レース群(軸wa50=257R/単指数90=335R)で
#   ファイル間の一致は独立な再現ではない。001は016以降の部分集合で、その差分
#   25〜35Rを『追加レース』として検証した。的中率の並びは引き継がれたが回収率は
#   引き継がれず、全条件の水準も概ね半分に落ちた(小標本)。そのため閾値近傍・
#   軸近傍・余裕的中数(損益分岐までの的中本数)で頑健性を判定して選び直した。
#   本線枠(実弾) 10条件 ─ 紐は全て『該当馬のうち紐馬指数最大の1頭』:
#     馬連 V1  軸[単指数90+&人気1以内]              × 単勝確率8.4          (◎ R174 的中25.9% 回収122.4%)
#     馬連 V2  軸[単指数90+&人気1以内&オッズ3以内]  × 複勝確率47.7&単指数50 (◎ R103 33.0% 124.6%)
#     馬連 V3  軸[単指数90+&人気1以内&オッズ3以内]  × 単勝確率8.4&単指数68  (◎ R120 28.3% 126.1%)
#     馬連 V4  軸[単指数90+&人気1以内&オッズ2以内]  × 単勝確率4.6          (◎ R104 29.8% 133.4%)
#     馬連 V5  軸[単指数90+&人気1以内&オッズ3以内]  × 複勝確率47.7&単指数68 (○ R71  35.2% 140.4%)
#     馬連 AR5 軸[単指数90+&人気1以内]              × 複勝確率47.7          (○ 維持)
#     馬連 AR6 軸[単指数90+&人気1以内]              × 勝率15.6&単指数68     (△ 維持・R56)
#     ワイド V6  軸[単指数90+&人気1以内&オッズ2以内] × 複勝確率47.7         (○ R71 64.8% 113.0% ・AR8置換)
#     ワイド V7  軸[wa50+&単指数90+&人気1以内&オッズ3以内] × 勝率4.6&複指数20 (○ R73 54.8% 120.8%)
#     ワイド AR7 軸[単指数90+&人気1以内]             × 勝率15.6&単指数68     (△ 維持・R56)
#   外した条件(定義は残置・GZ_MAIN_KEEP に戻せば復帰):
#     AR1/AR3/AR4 … 閾値を1段下げると回収率が大きく落ちる(観察扱い)
#     AR2 … 閾値・軸とも近傍が弱い / AR8 … 余裕0.9本 / AR9 … R40
#   ★馬単は本線枠なし。基準を通った馬単は全て騎手指数(騎手補正)を含み、
#     v136_001 の方針(騎手指数は期間バイアス疑いのため採用しない)に抵触するため。
#   ★馬連は同じ軸から最大7条件が発火しうる。同一ペアは1点に集約(先頭条件のタグ)。
#     相手が条件ごとに違う場合は1レースで複数点になる(本線の点数上限は従来どおり)。
#   ★新しい軸 'tan_idx>=90&pop<=1&odds<=2' / 'adj_win>=50&tan_idx>=90&pop<=1&odds<=3'
#     を追加(_axis_tan_idx に adj_win_min 引数を追加)。予想表上部の凡例は券種ごとに並べる。
#
# ★v136_005: 本線枠を『紐条件アーカイブ』(最適紐016〜026 再分析・馬単/馬連/ワイド
#   各3条件・ROI100%超を軸に的中率と再現性で選抜)の提案9条件へ全面差し替え。
#   次点本線枠(JITEN・勝ち筋マップ表示)は廃止した。
#   本線枠(実弾) 9条件:
#     AR1 馬単 軸[補正後勝率50%+]              × 複勝確率47.7以上            (◎)
#     AR2 馬単 軸[単指数90+]                    × 複勝確率47.7以上            (○)
#     AR3 馬単 軸[補正後勝率50%+]              × 複勝確率47.7以上&騎手補正48.1以上 (△要検証)
#     AR4 馬連 軸[補正後勝率50%+]              × 複勝確率47.7以上            (◎)
#     AR5 馬連 軸[単指数90+&人気1以内]         × 複勝確率47.7以上            (○)
#     AR6 馬連 軸[単指数90+&人気1以内]         × 勝率15.6&単指数68           (△要検証)
#     AR7 ワイド 軸[単指数90+&人気1以内]        × 勝率15.6&単指数68           (△要検証)
#     AR8 ワイド 軸[単指数90+&人気1以内&オッズ3以内] × 複勝確率47.7以上       (○)
#     AR9 ワイド 軸[補正後勝率50%+]            × 単指数68&騎手補正58.9以上   (△要検証)
#   ★◎○△は紐条件アーカイブ側の再現性評価(◎複数ファイルで100%超を一貫して確認/
#     ○大標本(R>=100)で確認/△単一ファイル・小標本(R<100)で要検証)であり、
#     このスクリプト側でCI下限を再算定したものではない。実弾は損失許容の範囲で。
#   ★軸・相手とも『補正後』の列(単勝確率/複勝確率/単指数)で判定する(最適紐008以来の
#     踏襲)。『騎手補正』は最適紐系の「騎手指数」に対応する表示列(KISHU_DEV_COL)で、
#     本コードの内部列『騎手指数』(生・部外秘)とは別物のため取り違えないこと。
#   ★人気・オッズは単勝人気/単勝オッズ(いずれも補正の概念がない生の列)で判定する。
#   ★旧 HA1/HA2/HB1/HB2(最適紐008提案4条件)は GZ_COND_DEFS に定義を残置。
#     GZ_MAIN_KEEP を戻せば復帰できる。
#   ★次点本線枠は JITEN_ENABLE = False で生成そのものを止めた(定義・ローダは残置)。
#     予想表(HTML/Excel/CSV)に【次点本線枠】の行は一切出力されなくなる。
# ★v136_003: ◎○の決定を『アンサンブル単勝確率』へ全面変更し、参考予想を刷新。
#   ・アンサンブル単勝確率 = 単勝オッズ由来確率(1/オッズの正規化)40%
#     + モデル予測単勝確率(補正後・表示列と同一)60%。脚質『追』の馬は減点
#     (ENSEMBLE_OIKOMI_PENALTY)。予想表の列には出さない(◎○決定専用)。
#   ・◎=アンサンブル単勝確率1位 / ○=同2位。
#   ・▲△☆(最大3頭)は『複勝確率(補正後)47.7%以上 または 紐馬指数60以上』の
#     馬(◎○を除く)に紐馬指数降順で付与。該当0頭のときのみ紐馬指数上位3頭。
#   ・単勝/複勝/本線枠/次点本線枠の買い目ロジック(軸選定は単勝確率_生ベース)は
#     変更していない。◎○▲△☆の表示と買い目の軸は必ずしも一致しない。
#   ・旧・参考枠(REF_WP1 / gz_reference.py 駆動)は全廃した。参考予想は
#     ◎ × ▲△☆の馬(上記と同一条件)への馬連・ワイドに一本化した
#     (詳細は _build_reference_himo 直上のコメントを参照)。
# ★v136_002: 参考枠を『軸=単勝確率1位』の自前ロジックへ全面変更(最適紐010 由来)。
#   ※本ロジックは v136_003 で REF_WP1 ごと廃止・置換済み(下記は変更履歴として残置)。
#   ・軸は常に単勝確率1位の馬(=◎)。相手は『紐馬指数60以上の馬のうち単指数が最大の1頭』1点。
#   ・券種は馬単(軸→相手)と馬連の2点。ワイドは全水準で回収率100%を割るため出さない。
#   ・相手の単指数の水準で採用条件(R1)と近似条件(R2〜R4)を段階表示する。
#     条件を満たす馬がいないレースは FB(検証条件ではない)として単指数1位へフォールバック。
# ★v136_001: 本線枠・次点本線枠を『最適紐008(278,100通り)再分析』の提案へ全面変更。
#   本線枠(実弾) 4条件:
#     HA1 ワイド 軸[補正後勝率40-49.9%] × 勝率15.6 & 単指数68  (N174 的中54.0% 回収104.1%)
#     HA2 馬連   同上                                          (N174 的中29.9% 回収111.3%)
#     HB1 馬連   軸[補正後勝率50%+]     × 複勝確率47.6          (N114 的中41.2% 回収118.5%)
#     HB2 馬単   同上                                          (N114 的中32.5% 回収139.7%)
#   次点本線枠(表示専用) 5条件: 下記 _JITEN_V136_RAW を参照。
#   ★主要な設計変更:
#     (1) 本線枠に『ワイド』を新設(従来は馬単/馬連のみ)。gz_wide_cond / wide_recs /
#         gz_stakes / 点数上限 / HTML・Excel・CSV・GUI の全経路にワイドを通した。
#     (2) 軸atomに補正後 単勝確率ベースの 'adj_win>=50'(50%以上) と
#         'adj_win40_49'(40〜49.9%の帯・50%以上とは排他) を追加。
#         最適紐008の探索は予想表の表示列(補正後)で行われているため、
#         本線枠も 生('_生')ではなく補正後列で判定する。
#     (3) 騎手指数を含む条件は採用しない。しきい値を上げても成績が伸びず
#         (41.2→48で的中低下)、対象レースが3割減ることから、騎手の実力ではなく
#         『騎手補正列を持つ新フォーマット期だけを残す日付フィルタ』として
#         効いている疑いが濃厚なため。検証は期間分割で行うこと。
#     (4) 券種は軸の強さで切り替える。軸が中程度(40-49.9%)はワイド+馬連、
#         軸が強い(50%+)は馬連+馬単。逆側の券種は回収率が100%を割る。
# ★v135_035: 次点本線枠に『おすすめ条件ランキング1～30位』(最適紐007)を追加。
#   JITEN_MAP_DEFS の先頭に30件を連結し、既存の次点条件と同じ枠で買い目を生成する。
#   ★判定は補正後の列(単勝確率/複勝確率)で行う。軸『単勝確率50%以上』は補正後勝率軸
#     'wa50'、『単指数90以上』は単指数のみ軸 'ti90' を _jiten_axis に追加。表示は
#     順位・タイプ・回収率/的中率/N を注記(定義は _OSUSUME_RANK_RAW を参照)。
# ★v135_021: 本線枠に P1/P2(馬単・軸wp50+×複勝率40系)を追加。紐カバー枠に
#   R3(馬連・軸騎手50×単勝率40&オッズ1-10)を少額で追加。F4/S2/U1は残置。
#   出自: prob_model_conditions_with_axis.json を Benjamini-Hochberg(FDR)で
#   再判定し、一貫性・大標本・的中数・クリーン構造で選抜(FDR有意0件のグレー採用)。
#   採用候補のうち U1同一/U1部分集合の重複2件は不採用。詳細は各条件直上コメント。
# ★v135_013: 本線枠に F4(馬単・S2∩期待値上位2頭) を追加。S2の上位互換候補で
#   検証CI下限112.6%。ただし選択バイアスが残るため S2 と並行記録して比較する。
#   F4は f4_params.json が必要(fit_f4_params.py が生成)。無い場合は発火しない。
# ★v135_012: 本線枠に S2(馬単・軸wp40+×複勝確率上位2頭) を追加。
#   S2のみ出自が異なり、入れ子ローリング検証(13か月/3,802R・選択汚染ゼロ)で
#   ブロックブートストラップCI下限 103.3% を確認した唯一の条件。詳細は
#   GZ_COND_DEFS の 'S2' 直上コメントを参照。
# ★買い目条件は prob_model_conditions_with_axis.json の robust条件から選抜した
#   本線: 馬単(F4/S2/U1/U2/U3/U7/U8) 馬連(R1)。紐カバー枠は単勝確率軸テンプレ。
#   ★v135_011: 各条件に tier を付与し『本線枠(main)』と『紐カバー枠(ana)』に分離。
#     本線枠 = 1点あたり的中率 p>=0.10 かつ ROI下限(近似) lo>=100 の条件。
#              U7/U1/R1/U3/U2/U8
#     紐カバー枠 = 上記を満たさない条件(低的中率の高配当依存、または下限不足)。
#              R2/U4/U6/R4/U5/R3
#     枠ごとに点数上限(財布)を分けており、穴条件が本線条件の枠を食わない。
#     「有力」表示・Excel/CSV/HTMLの表も枠ごとに分けて出力する。
#   HTML/Excel/CSV/GUI のラベル・配色はすべてそこから自動生成する。
#   詳細は GZ_COND_DEFS 直上のコメントを参照。
#
# ★賭け金は均等買い(1点 GZ_FLAT_UNIT 円)+1レース点数上限。
#   検証ROIが1点100円の均等買いで算出された値なので、それに合わせている。
#
# ★単勝/複勝は「安定運用」(_stable_axis_ok)系のみ。会場別に検証した
#   安定条件を満たす◎に限って1点提示する。
#
# ★出力: CSV / カラフルExcel / HTML予想表 / PDF(HTMLから変換) / tkinter GUI。
#
# ★ゲート(買い目は出すが備考を『参考』に強制するもの)
#   ・開催地ゲート : NSL_REFERENCE_ONLY_VENUES
#   ・信頼度区分ゲート: 実配当検証で合計回収が安定赤字だった
#     『非常に高い/高い/非常に低い』区分のレース。
#
# ★常に空で出力されないもの(構造は残置): 複勝(通常)/穴複勝/ワイド/三連複。
#   いずれも検証で優位が確認できず廃止。表示側は空判定でスキップされる。
#
# ※採用条件はいずれもCI下限(近似)が100%を明確に上回るわけではなく、
#   利益が統計的に確認された条件ではない。損失許容の範囲で運用すること。
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

# ─────────────────────────────────────────────────────────────
# 単勝/複勝確率 LightGBMモデル (keiba_prob_model) の読み込み
#   ・generate_analysis 内の確率算出をこのモデル出力で完全置換する。
#   ・モデル成果物(win_model.txt / place_model.txt / prob_model_meta.json)は
#     既定でこのスクリプトと同じフォルダを探索。環境変数 PROB_MODEL_DIR で上書き可。
#   ・モデルは一度だけロードしてキャッシュする。
# ─────────────────────────────────────────────────────────────
_PROB_MODEL_CACHE = None

# ════════════════════════════════════════════════════════════════
# ★ 推論時 horselist 結合 (v133_modified007)
#   学習モデルは horselist(血統/成績/タイム/月齢/頻度エンコード)由来の
#   27特徴量を使う。推論時にこれを渡さないと全馬-1埋めになり予測が壊れる。
#   → 予想対象CSVと同じフォルダの *_horselist.csv を読み、
#     (競馬場, 競走年月日, レース番号, 馬番) で各馬に 'hl' を結合する。
#   ファイル名は日付で変わるが '*_horselist.csv' の glob で自動追従。
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
# ★v136_003: アンサンブル単勝確率(◎◯決定専用。予想表には表示しない)
#   ・単勝オッズ由来の確率(1/オッズをレース内で正規化)と、
#     モデル予測の単勝確率(補正後・表示列と同一)を加重平均する。
#   ・脚質(展開列)が『追』の馬は、道中不利になりやすくオッズ/指数が
#     過大評価されがちなため、アンサンブル単勝確率のみ減点する。
#   ・本線枠/次点本線枠/単勝/複勝の買い目ロジックはこの確率を使わない
#     (従来どおり 単勝確率_生 / 単勝確率(補正後) を使用、変更なし)。
#     アンサンブル単勝確率は ◎(1位)/○(2位) の決定にのみ用いる。
# ══════════════════════════════════════════════════════════════════
ENSEMBLE_ODDS_WEIGHT    = _gz_env_num('ENSEMBLE_ODDS_WEIGHT', 0.40, float)   # オッズ由来確率の比率
ENSEMBLE_MODEL_WEIGHT   = _gz_env_num('ENSEMBLE_MODEL_WEIGHT', 0.60, float)  # モデル予測確率の比率
ENSEMBLE_OIKOMI_PENALTY = _gz_env_num('ENSEMBLE_OIKOMI_PENALTY', 0.90, float)  # 脚質『追』の減点係数

# ★v136_003: 紐馬(◎◯以外の予想印 ▲△☆ / 参考予想の相手)の採用条件。
#   複勝確率(補正後)47.7%以上 または 紐馬指数60以上。該当馬が0頭のときのみ
#   紐馬指数上位3頭にフォールバックする。
HIMO_PLACE_PROB_MIN = _gz_env_num('HIMO_PLACE_PROB_MIN', 47.7, float)
HIMO_HIMOBA_IDX_MIN = _gz_env_num('HIMO_HIMOBA_IDX_MIN', 60.0, float)

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


_NUMERIC_RACECARD_COLS = ['単指数', '複指数', '騎手指数', '単勝人気',
                          '単勝オッズ', '番', '斤量']


def to_numeric_racecard(df):
    """出馬表の数値列を安全に数値化して返す(欠損は0補完)。

    ★v135_010 バグ修正: 全角数字('１','２.5')は pd.to_numeric が NaN にするため、
      fillna(0) と組み合わさって『全馬の馬番が0』『人気が全部0』のように
      無警告でデータが潰れていた(◎=0番、買い目が 0→0 になり実用不能)。
      レース番号列は以前から半角化していた(全角'１０R'の実例あり)ので、
      同じCSVの数値列にも全角が来る前提で、数値化の前に必ず半角化する。
    """
    for c in [c for c in _NUMERIC_RACECARD_COLS if c in df.columns]:
        col = df[c]
        # ★数値でない列(=文字列として読まれた列)だけ半角化する。
        #   pandas 3.0 では文字列列の dtype が object ではなく str になるため、
        #   dtype==object の判定では取りこぼす。is_numeric_dtype で判定する。
        if not pd.api.types.is_numeric_dtype(col):
            col = col.astype(str).str.translate(_Z2H_DIGITS).str.strip()
        df[c] = pd.to_numeric(col, errors='coerce').fillna(0)
    return df


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

def _get_horselist_table(folder):
    """フォルダ内の *_horselist.csv を読み、(venue,date,race_no,umaban)->行dict。"""
    global _HL_MISSING_WARNED
    if not folder:
        return {}
    folder = str(folder)
    # ★v136_021: 旧版はフォルダ名だけでキャッシュしていたため、起動後に horselist を置いた/更新した
    #   (騎手変更など)場合も最初の内容を使い続けた。ファイルの状態を含めてキャッシュする。
    try:
        import glob as _glob_hl
        _sig = tuple(sorted((_p, _os_gz.path.getmtime(_p), _os_gz.path.getsize(_p))
                            for _p in _glob_hl.glob(_os_gz.path.join(folder, '*_horselist.csv'))))
    except Exception:
        _sig = ()
    _ck = (folder, _sig)
    if _ck in _HORSELIST_TABLE_CACHE:
        return _HORSELIST_TABLE_CACHE[_ck]
    for _k in [k for k in _HORSELIST_TABLE_CACHE if isinstance(k, tuple) and k[0] == folder]:
        del _HORSELIST_TABLE_CACHE[_k]
    table = {}
    _err = None
    try:
        import horselist_features as _HLmod
        table = _HLmod.load_horselist(folder, verbose=False)
    except Exception as _e:
        _err = _e
        table = {}
    if not table and not _HL_MISSING_WARNED:
        # ★無警告で全馬-1埋め(精度低下)になるのを防ぐため必ず知らせる
        if _err is not None:
            print(f'  [horselist警告] horselist_features の読込に失敗: {_err}\n'
                  f'    horselist_features.py を本スクリプトと同じフォルダに置いてください。\n'
                  f'    このままでも動作しますが、血統/成績等の特徴が欠損し予測精度が低下します。')
        else:
            print(f'  [horselist警告] {folder} に *_horselist.csv が見つかりません。\n'
                  f'    予想対象CSVと同じフォルダに当日の horselist を置いてください。\n'
                  f'    このままでも動作しますが、血統/成績等の特徴が欠損し予測精度が低下します。')
        _HL_MISSING_WARNED = True
    _HORSELIST_TABLE_CACHE[_ck] = table
    return table

def _attach_hl_to_group(group, folder):
    """group(1レース分)の各行に 'hl'(horselist行dict) を付けて返す。
    horselistが無い/未マッチの馬は hl=None(学習時と同じ欠損-1扱い)。"""
    table = _get_horselist_table(folder)
    hl_list = [None] * len(group)
    if not table:
        return hl_list
    # レース識別: 競馬場='場所'列 / レース番号='レース'(全角対応) / 馬番='番'。
    # ★対象日の決定(複数日分の horselist がある場合の誤マッチ防止):
    #   優先1: レースカードのファイル名から推定した日付(CURRENT_RACECARD_DATE)
    #   優先2: horselist の日付が1種類だけならその日
    #   どちらも決まらなければ日付フォールバックはせず、厳密キーのみ試す。
    _dates = set(k[1] for k in table.keys())
    _single_date = None
    if CURRENT_RACECARD_DATE is not None and CURRENT_RACECARD_DATE in _dates:
        _single_date = CURRENT_RACECARD_DATE
    elif len(_dates) == 1:
        _single_date = next(iter(_dates))
    # 複数日ありで対象日が特定できない場合の警告(1回だけ)
    if _single_date is None and len(_dates) > 1:
        global _HL_MULTIDATE_WARNED
        if not _HL_MULTIDATE_WARNED:
            print('  [horselist警告] 複数日分の *_horselist.csv があり対象日を'
                  '特定できません。レースカード名に日付(例 0712 / 20260712)を'
                  '含めるか、対象日のhorselistのみをフォルダに置いてください。'
                  '日付をまたぐ誤結合を避けるため、この状態では horselist を'
                  '結合しません(予測精度が低下します)。')
            _HL_MULTIDATE_WARNED = True
    def _venue_of():
        try:
            return _norm_venue_hl(group['場所'].iloc[0]) if '場所' in group.columns else ''
        except Exception:
            return ''
    def _rno_of():
        for _c in ('レース', 'レース数', 'レース番号', 'R'):
            if _c in group.columns:
                _n = _to_int_z2h(group[_c].iloc[0])   # 全角'１'にも対応
                if _n is not None:
                    return _n
        return None
    venue = _venue_of()
    rno = _rno_of()
    for _pos, (_, r) in enumerate(group.iterrows()):
        num = _to_int_z2h(r.get('番'))   # 全角馬番にも対応
        if num is None:
            continue
        matched = None
        if _single_date is not None and rno is not None:
            # 対象日が確定 → 厳密キー(日付込み)でマッチ
            matched = table.get((venue, _single_date, rno, num))
            if matched is None:
                # 会場名の表記ゆれ吸収のため、日付固定のまま会場総当たり
                for k, row in table.items():
                    if (k[1] == _single_date and k[2] == rno and k[3] == num
                            and _norm_venue_hl(k[0]) == venue):
                        matched = row
                        break
        elif len(_dates) == 1 and rno is not None:
            # 日付が1種類しか無い(=単一日) → その日で総当たり
            _only = next(iter(_dates))
            matched = table.get((venue, _only, rno, num))
        # ★複数日ありで対象日不明のときは日付をまたぐ誤結合を避けるため
        #   フォールバックしない(matched=None のまま=学習時と同じ欠損扱い)。
        hl_list[_pos] = matched
    return hl_list


# ─────────────────────────────────────────────────────────────
# 出馬表CSVの読込ヘルパー(学習コード keiba_prob_model と同じ方針)
#   ・エンコーディングは cp932/utf-8-sig/utf-8/shift_jis を自動判定し、
#     必須列(場所/番)が揃った読み方を採用する。utf-8-sig 決め打ちだと
#     cp932 のCSVで文字化け or 読込失敗するため。
#   ・列名の空白除去＋別名統一(予想 オッズ→単勝オッズ, 人気→単勝人気)。
#     一部の日のファイルは列名が異なるため、ここで正規名へ寄せる。
#   ※ ここでの単勝オッズ/単勝人気は市場オッズではなく別AIの予測値。
#     学習時に特徴量として使うため、推論でも同じ列を読み込む必要がある。
# ─────────────────────────────────────────────────────────────
_COL_ALIASES_YOSOU = {
    '単勝オッズ': ('単勝オッズ', '予想オッズ', 'オッズ', '単オッズ'),
    '単勝人気': ('単勝人気', '人気', '単人気', '予想人気'),
}


def _normalize_racecard_columns_yosou(df):
    import re as _re2
    stripped = {c: _re2.sub(r'\s+', '', str(c)) for c in df.columns}
    df = df.rename(columns=stripped)
    for canon, aliases in _COL_ALIASES_YOSOU.items():
        if canon in df.columns:
            continue
        for a in aliases:
            a2 = _re2.sub(r'\s+', '', a)
            if a2 in df.columns:
                df = df.rename(columns={a2: canon})
                break
    return df


def read_racecard_csv(path, require_cols=('場所', '番')):
    """出馬表CSVを自動エンコーディング＋列名正規化して読む。"""
    # ★ horselist 結合用: 読み込むCSVのフォルダを記録(同フォルダの
    #   *_horselist.csv を推論時に自動参照する)。日付でファイル名が
    #   変わっても glob '*_horselist.csv' で追従する。
    global CURRENT_RACECARD_DIR, CURRENT_RACECARD_DATE
    try:
        import os as _os_rc
        CURRENT_RACECARD_DIR = _os_rc.path.dirname(_os_rc.path.abspath(str(path)))
        CURRENT_RACECARD_DATE = _guess_date_from_name(path)
    except Exception:
        pass
    last = None
    first_ok = None
    for enc in ('cp932', 'utf-8-sig', 'utf-8', 'shift_jis'):
        try:
            df = pd.read_csv(path, encoding=enc)
            df.columns = [str(c).lstrip('\ufeff') for c in df.columns]
        except Exception as e:
            last = e
            continue
        df = _normalize_racecard_columns_yosou(df)
        if first_ok is None:
            first_ok = df
        if all(c in df.columns for c in require_cols):
            return df
    if first_ok is not None:
        return first_ok
    raise last


def _find_prob_model_dir():
    cand = []
    env = _os.environ.get('PROB_MODEL_DIR')
    if env:
        cand.append(env)
    here = _os.path.dirname(_os.path.abspath(__file__))
    cand.extend([here, _os.path.join(here, 'models'), '.', './models'])
    for d in cand:
        if d and _os.path.exists(_os.path.join(d, 'prob_model_meta.json')):
            return d
    return None

def _get_prob_model():
    """ProbModel をロードして返す(キャッシュ)。見つからなければ明示エラー。"""
    global _PROB_MODEL_CACHE
    if _PROB_MODEL_CACHE is not None:
        return _PROB_MODEL_CACHE
    try:
        from keiba_prob_model import ProbModel
    except Exception as e:
        raise RuntimeError(
            'keiba_prob_model.py が import できません。'
            '同じフォルダに配置してください。詳細: %s' % e)
    mdir = _find_prob_model_dir()
    if mdir is None:
        raise FileNotFoundError(
            '確率モデル(prob_model_meta.json 等)が見つかりません。'
            'keiba_prob_model.py で学習し、成果物をこのスクリプトと同じフォルダ'
            'または環境変数 PROB_MODEL_DIR のフォルダに置いてください。')
    _PROB_MODEL_CACHE = ProbModel.load(mdir)
    return _PROB_MODEL_CACHE

# ─────────────────────────────────────────────────────────────
# ★ 投資判定（analysis_summary.json 由来の回収率プラス条件）
#   軸②(穴馬軸推奨) / 軸③(単指数90+&1番人気&複勝確率80+) について
#   回収率100%以上の券種の買い目 備考欄に「投資」を付与する。
# ─────────────────────────────────────────────────────────────
INVEST_ROI_MIN = 100.0   # この回収率(%)以上を「投資」対象とする


def _load_invest_conditions(json_path: 'Path') -> dict:
    """
    analysis_summary.json を読み、軸②/軸③ごとに回収率100%以上の
    「投資」対象券種を抽出する。

    Returns:
      {
        'hole':   {'tansho':bool,'fukusho':bool,'wide_metric':str|None},  # 軸②
        'honmei': {'tansho':bool,'fukusho':bool,'wide_metric':str|None},  # 軸③
        'loaded': bool,
      }
    """
    result = {
        'hole':   {'tansho': False, 'fukusho': False, 'wide_metric': None},
        'honmei': {'tansho': False, 'fukusho': False, 'wide_metric': None},
        'loaded': False,
    }
    try:
        with open(json_path, encoding='utf-8') as f:
            data = _json.load(f)
    except Exception:
        return result

    def _fill(section_key, dst):
        sec = data.get(section_key)
        if not isinstance(sec, dict):
            return
        if float(sec.get('axis_tansho_recovery', 0) or 0) >= INVEST_ROI_MIN:
            dst['tansho'] = True
        if float(sec.get('axis_fukusho_recovery', 0) or 0) >= INVEST_ROI_MIN:
            dst['fukusho'] = True
        # ワイドは「最良ワイド回収率」が100%以上なら、その紐選定指標を採用
        if float(sec.get('best_wide_recovery', 0) or 0) >= INVEST_ROI_MIN:
            dst['wide_metric'] = sec.get('best_wide_recovery_metric')

    _fill('anaba_himo', result['hole'])    # 軸②: 穴馬軸推奨
    _fill('himo3',      result['honmei'])  # 軸③: 本命強軸
    result['loaded'] = True
    return result


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
    # ★ 投資判定に使う集計サマリー(回収率プラス条件の根拠)
    INVEST_SUMMARY_PATH  = Path(_os.environ.get(
        'RACE_INVEST_SUMMARY', 'c:/Users/pino/race_prediction/tihoukeiba002/analysis_summary.json'))

    WIDE_ODDS_DIVISOR    = 3.15
    UMATAN_RATIO         = 1.85
    UMAREN_RATIO         = 3.08
    # ★ v131.1: 複勝確率ベースのワイド予測オッズ用係数。
    #   WIDE_TAKEOUT   = 地方競馬ワイドの概算払戻率(控除率≒0.20-0.25 → 払戻0.75-0.80)。
    #   WIDE_JOINT_CORR= 強い2頭が同時に3着内へ入る際の負相関(共存しにくさ)を織り込む
    #                    同時確率の割引係数(<1 で joint をやや下げ=オッズをやや上げる)。
    WIDE_TAKEOUT         = 0.78
    WIDE_JOINT_CORR      = 0.90
    # ★ 騎手指数によるオッズ補正(v015)。
    #   人気騎手(=レース平均より騎手指数が高い馬)のオッズを下げ、
    #   非人気騎手(平均より低い馬)のオッズを上げる。基準はレース内相対。
    #   係数 = (騎手指数 / レース平均)^(-STRENGTH)。
    #     STRENGTH=0 で無補正、値を上げるほど補正が強い。
    #   補正後は adjusted_odds 昇順で adjusted_popularity を付け直す。
    JOCKEY_ODDS_ADJ_ENABLE   = True
    JOCKEY_ODDS_ADJ_STRENGTH = 0.30   # 補正の強さ(推奨 0.2-0.5)
    JOCKEY_ODDS_ADJ_MIN      = 0.70   # 1頭あたり係数の下限(下げ過ぎ防止)
    JOCKEY_ODDS_ADJ_MAX      = 1.50   # 1頭あたり係数の上限(上げ過ぎ防止)

    # ★v135_005: 表示専用 補正騎手指数のモード ('resid' or 'ratio')
    #   買い目条件では使わない(条件は生の騎手指数)。
    KISHU_DEV_MODE = 'resid'


# ★ 投資条件はファイル更新に追従できるよう遅延ロード＆キャッシュ
_INVEST_COND_CACHE = None
def get_invest_conditions():
    global _INVEST_COND_CACHE
    if _INVEST_COND_CACHE is None:
        _INVEST_COND_CACHE = _load_invest_conditions(Config.INVEST_SUMMARY_PATH)
    return _INVEST_COND_CACHE


def _parse_args():
    parser = argparse.ArgumentParser(description='新地方競馬予想 v078 GUI版')
    parser.add_argument('--data',   type=Path, default=Config.DEFAULT_DATA_PATH)
    parser.add_argument('--output', type=Path, default=Config.DEFAULT_OUTPUT_DIR)
    parser.add_argument('--deba', type=Path, default=None,
                        help='★v136_019 出馬表HTMLフォルダ(起動時に読み込む)')
    args, _ = parser.parse_known_args()
    return args

_args = None


# ============================================================
# ★ 軸グレード式(バックテスト最適化結果 axis_grade_formula)
#   実バックテストで S>A>B>C>D が3着内率で単調降順になるよう最適化した
#   標準化線形スコア式。旧 _base_grade(複合スコア+gap閾値)+荒れ度補正は
#   成績と非単調(B>A 等)だったため、これに置き換える。
#   既定は埋め込み定数。スクリプトと同じフォルダ(または環境変数
#   AXIS_GRADE_FORMULA_PATH)に axis_grade_formula.json があれば自動優先採用。
# ============================================================
_AGF_FEATS = ['複合スコア', 'gap', '複勝確率', '単勝確率', '単勝期待値', '紐馬指数']
_AGF_MEAN = {'複合スコア': 79.085781, 'gap': 9.391883, '複勝確率': 71.865032,
             '単勝確率': 41.368702, '単勝期待値': 121.082607, '紐馬指数': 79.947425}
_AGF_STD = {'複合スコア': 6.035827, 'gap': 6.598482, '複勝確率': 17.025765,
            '単勝確率': 15.602336, '単勝期待値': 42.190823, '紐馬指数': 10.335227}
_AGF_W = {'複合スコア': 0.182347, 'gap': -0.015885, '複勝確率': 0.01457,
          '単勝確率': 0.580586, '単勝期待値': -0.789039, '紐馬指数': -0.081377}
_AGF_CUTOFFS = [['S級軸', 0.8725429229492219], ['A級軸', 0.2270214332179692],
                ['B級軸', -0.30472027387484824], ['C級軸', -0.8523026817108669], ['D級軸', None]]
_AGF_SRC = 'embedded'

def _load_axis_grade_formula():
    """外部 axis_grade_formula.json があれば定数を差し替える(無ければ埋め込み維持)。"""
    import os as _os3
    global _AGF_FEATS, _AGF_MEAN, _AGF_STD, _AGF_W, _AGF_CUTOFFS, _AGF_SRC
    cands = []
    envp = _os3.environ.get('AXIS_GRADE_FORMULA_PATH')
    if envp:
        cands.append(envp)
    try:
        cands.append(_os3.path.join(_os3.path.dirname(_os3.path.abspath(__file__)), 'axis_grade_formula.json'))
    except Exception:
        pass
    cands.append('axis_grade_formula.json')
    for p in cands:
        try:
            if p and _os3.path.exists(p):
                with open(p, encoding='utf-8') as f:
                    d = _json.load(f)
                _AGF_FEATS = list(d['features'])
                _AGF_MEAN = dict(d['standardize']['mean']); _AGF_STD = dict(d['standardize']['std'])
                _AGF_W = dict(d['weights'])
                _AGF_CUTOFFS = [[g, (None if c is None else float(c))] for g, c in d['grade_cutoffs']]
                _AGF_SRC = p
                return p
        except Exception:
            continue
    return None
_load_axis_grade_formula()


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

# ブレンド勝率: p ∝ (単勝確率)^A × (補正後オッズ由来確率)^B をレース内で正規化
JUDGE_BLEND_A = 0.4317
JUDGE_BLEND_B = 0.7677
JUDGE_ODDS_MISSING = 199.0      # オッズ欠損馬は最低人気相当として扱う
# 推定3着内率: logit(p3) = SLOPE × logit(Harville3着内確率) + INTERCEPT
JUDGE_TOP3_SLOPE = 0.6374
JUDGE_TOP3_INTERCEPT = -0.1356
# 軸級: ◎の推定勝率(%)の下限。これ未満は D級
#   実績(7/16〜9/19 2,324R): S 勝率62.0%/複勝圏88.1%, A 42.6/77.6, B 31.2/67.0, C 27.8/62.4, D 21.1/49.9
JUDGE_GRADE_CUTS = [('S', 50.0), ('A', 40.0), ('B', 31.0), ('C', 24.0)]
# 信頼度: 数値 = ◎の推定3着内率(%)。区分境界(色は旧区分と同じ)
#   実績(7/16〜9/19 2,324R): 非常に高い 88.6%, 高い 78.5, 中程度 68.2, 低い 59.4, 非常に低い 48.5
JUDGE_CONF_TIERS = [
    (83, '非常に高い', 'FFFFFF', '375623', 'C6EFCE'),
    (74, '高い',       'FFFFFF', '1F4E79', 'BDD7EE'),
    (64, '中程度',     '000000', 'FFEB9C', 'FFF2CC'),
    (55, '低い',       'FFFFFF', 'ED7D31', 'F8CBAD'),
    (-1, '非常に低い', 'FFFFFF', 'C00000', 'FFC7CE'),
]
# レース分類(上から順に判定)
JUDGE_CAT_INVEST_WIN = 60.0     # 投資: 推定勝率60%以上 かつ
JUDGE_CAT_INVEST_TOP3 = 88.0    #       推定3着内率88%以上
JUDGE_CAT_STRONG_TOP3 = 76.0    # 有力: 推定3着内率76%以上
# ★v136_009: 2頭軸の判定は停止。9月(未使用データ)で ◎-対抗 馬連的中 1/20R(予測35%)、
#   7/16〜9/19 通算でも 26%(85R)と有力レース(約30%)を下回り、区分として機能しなかった。
#   該当していたレースは 標準レース(推定3着内率55%未満なら見送り=表示は標準) になる。
JUDGE_CAT_TWO_ENABLE = False    # True で再開(下の2つの閾値を使用)
JUDGE_CAT_TWO_Q = 30.0          # 2頭軸: ◎-対抗の馬連確率30%以上 かつ
JUDGE_CAT_TWO_RIVAL = 25.0      #        対抗のブレンド勝率25%以上
JUDGE_CAT_SKIP_TOP3 = 55.0      # 見送り: 推定3着内率55%未満(表示は _RACE_CAT_MASK で標準レース)


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
BADO_VERSION = 'v136_021'


def bado_version_meta():
    """予想表HTMLの <head> に入れる版情報の meta タグ(2本)を返す。"""
    _logic = 'judge=%s;ref=%s;honsen=sheet1;stat=%s;ura=50;obs=st1-6;dq=1' % ('new' if USE_NEW_JUDGE else 'old',
                                                       'new' if REF_NEW_LOGIC else 'old',
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
# 本線/次点/観察の条件コード → 軸の種類
STAT_AXIS_OF = {'UR1': 'UR1', 'OW1': 'UR1', 'UR2': 'UR2', 'WD1': 'WD1', 'OU1': 'OU1'}
STAT_AXIS_JP = {'UR1': 'UR1軸(単指数90・1.9倍以下)', 'UR2': 'UR2軸(単指数90-93・1.4倍以下)',
                'WD1': 'WD1軸(単確50・単指数90・1人気)', 'OU1': 'OU1軸(単確50・1.9倍以下)'}
# ★v136_020: 観察条件(統計1位を紐・表示と記録のみ・実弾外)。値は%。
#   p_min=紐の統計勝率 / ax_min=軸の統計勝率 / fp_min=紐の複勝確率(表示値) /
#   ratio_min=紐の統計勝率 ÷ 単勝確率(レース内で合計100に正規化)
STAT_OBS_DEFS = [
    dict(tag='ST1', axis='UR1', kind='umaren', p_min=10.0, ax_min=40.0,
         label='紐統計10%以上&軸統計40%以上', n=35, hit=57.1, roi=145.1),
    dict(tag='ST2', axis='UR1', kind='wide', fp_min=40.0, ratio_min=1.0,
         label='紐複勝40以上&統計≧単確', n=35, hit=82.9, roi=117.7),
    dict(tag='ST3', axis='WD1', kind='umaren', fp_min=40.0, p_min=10.0,
         label='紐複勝40以上&紐統計10%以上', n=32, hit=56.2, roi=161.2),
    dict(tag='ST4', axis='WD1', kind='wide', fp_min=45.0, p_min=15.0,
         label='紐複勝45以上&紐統計15%以上', n=22, hit=95.5, roi=129.1),
    dict(tag='ST5', axis='OU1', kind='umaren', p_min=15.0, ax_min=50.0,
         label='紐統計15%以上&軸統計50%以上', n=30, hit=70.0, roi=142.0),
    dict(tag='ST6', axis='OU1', kind='wide', p_min=15.0, ax_min=50.0,
         label='紐統計15%以上&軸統計50%以上', n=30, hit=90.0, roi=115.7),
]
for _d in STAT_OBS_DEFS:
    STAT_AXIS_OF[_d['tag']] = _d['axis']


def _stat_ura_of(analysis, code):
    """条件コード(UR1/UR2/WD1/OW1/OU1/ST*)の軸に統計裏付けがあれば {ban, ps, ok} を返す。"""
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
JOCKEY_DAY_IDX = {}          # (場, 騎手) -> [騎手指数...]  予想CSV1日分
RUN_ERRORS = []              # _run_output 中に分析できなかったレース
_JK_VARIANTS = str.maketrans({'濱': '浜', '髙': '高', '﨑': '崎', '邊': '辺', '邉': '辺', '齋': '斎', '齊': '斉',
                              '澤': '沢', '廣': '広', '德': '徳', '眞': '真', '惠': '恵', '國': '国', '櫻': '桜',
                              '龍': '竜', '嶋': '島', '嶌': '島', '冨': '富', '瀨': '瀬'})


def _jk_nz(s):
    s = unicodedata.normalize('NFKC', str(s or ''))
    s = re.sub(r'[\s▲△☆◇★*※]', '', s)
    return s.translate(_JK_VARIANTS)


def _jockey_same(a, b):
    """予想CSVとhorselistで騎手の略し方が違う(例 小笠原/小笠羚・多田羅/多田誠)ので、先頭2文字が同じなら同一騎手。"""
    na, nb = _jk_nz(a), _jk_nz(b)
    if not na or not nb or na == nb:
        return True
    return len(na) >= 2 and len(nb) >= 2 and na[:2] == nb[:2]


def _jockey_day_build(df):
    JOCKEY_DAY_IDX.clear()
    try:
        if '騎手' not in df.columns or '騎手指数' not in df.columns:
            return
        for v, j, x in zip(df['場所'], df['騎手'], pd.to_numeric(df['騎手指数'], errors='coerce')):
            if x == x:
                JOCKEY_DAY_IDX.setdefault((_norm_venue_hl(v), _jk_nz(j)), []).append(float(x))
    except Exception as _e:
        print(f'  [騎手変更] 騎手指数表の作成に失敗: {_e}')


def _jockey_day_idx(venue, name):
    """同じ日・同じ場の予想CSVにある騎手指数の中央値(新騎手の値)。見つからない/候補が複数なら None。"""
    n = _jk_nz(name)
    xs = JOCKEY_DAY_IDX.get((venue, n))
    if not xs:
        cand = [k for k in JOCKEY_DAY_IDX if k[0] == venue and _jockey_same(k[1], n)]
        if len(cand) != 1:
            return None
        xs = JOCKEY_DAY_IDX[cand[0]]
    return float(np.median(xs))


def _apply_jockey_changes(group):
    """horselist の騎手名が予想CSVと違う馬を新騎手に置き換える(group をその場で変更)。変更の一覧を返す。"""
    out = []
    if '騎手' not in group.columns:
        return out
    try:
        hl = _attach_hl_to_group(group, CURRENT_RACECARD_DIR)
        venue = _norm_venue_hl(group['場所'].iloc[0]) if '場所' in group.columns else ''
    except Exception:
        return out
    for pos, idx in enumerate(group.index):
        h = hl[pos] if pos < len(hl) else None
        if not h:
            continue
        new = str(h.get('騎手名') or '').strip()
        old = str(group.at[idx, '騎手'] or '').strip()
        if not new or new.lower() == 'nan' or not old or _jockey_same(old, new):
            continue
        new_disp = re.sub(r'\s', '', unicodedata.normalize('NFKC', new))
        old_idx = group.at[idx, '騎手指数'] if '騎手指数' in group.columns else None
        new_idx = _jockey_day_idx(venue, new_disp)
        if new_idx is not None and '騎手指数' in group.columns:
            # 騎手指数列が int64 だと中央値(69.5 等)を代入できず pandas 3 で TypeError になり、
            #   騎手名だけ『※』に変わって指数は旧騎手のまま・変更一覧も空、という半端な状態になっていた。
            if not pd.api.types.is_float_dtype(group['騎手指数']):
                group['騎手指数'] = pd.to_numeric(group['騎手指数'], errors='coerce').astype(float)
            group.at[idx, '騎手指数'] = new_idx
        group.at[idx, '騎手'] = new_disp + JK_CHANGE_MARK
        out.append(dict(ban=_to_int_z2h(group.at[idx, '番']), name=str(group.at[idx, '馬名']),
                        old=old, new=new_disp,
                        idx_old=(float(old_idx) if old_idx is not None and old_idx == old_idx else None),
                        idx_new=new_idx))
    if out:
        print('  [騎手変更] ' + ' / '.join('%s番%s %s→%s' % (c['ban'], c['name'], c['old'], c['new']) for c in out))
    return out


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
        if c.get('idx_new') is not None:
            _ix = '騎手指数%s→%.0f' % ('-' if c.get('idx_old') is None else '%.0f' % c['idx_old'], c['idx_new'])
        else:
            _ix = '騎手指数は旧騎手のまま(新騎手が当日CSVに無い)'
        out.append('騎手変更 %s番%s %s→%s(horselist)・%s' % (c['ban'], c['name'], c['old'], c['new'], _ix))
    _n, _m = analysis.get('hl_n') or 0, analysis.get('hl_m') or 0
    if _n and _m / _n < HL_RACE_WARN_RATE:
        out.append('horselist未結合 %d/%d頭 → 単勝確率が不正確(当日のhorselistを確認)' % (_m, _n))
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
    return dict(tot=tot, m=m, rate=rate, low=low, jk=jk, scr=scr, errors=errs, n_failed=n_failed,
                severe=bool((rate is not None and rate < HL_WARN_RATE) or errs or n_failed))


def data_quality_lines(q):
    out = []
    if q['rate'] is not None and q['rate'] < HL_WARN_RATE:
        out.append('【重要】horselist結合率 %.0f%%(%d/%d頭)。当日の *_horselist.csv が予想CSVと同じフォルダに無いか、'
                   '日付が合っていません。単勝確率・買い目が不正確です。' % (q['rate'] * 100, q['m'], q['tot']))
        if q['low']:
            out.append('  未結合のレース: ' + ' '.join(q['low'][:12]) + (' ほか' if len(q['low']) > 12 else ''))
    elif q['low']:
        out.append('horselist未結合のレース: ' + ' '.join(q['low'][:12]))
    if q['errors'] or q['n_failed']:
        out.append('【重要】分析に失敗したレース %d件(予想表に出ていません):' % max(len(q['errors']), q['n_failed']))
        out.extend('  ' + e for e in q['errors'][:8])
    if q['jk']:
        out.append('騎手変更 %d件(horselistの騎手で計算):' % len(q['jk']))
        out.extend('  %s %s番%s %s→%s%s' % (r, c['ban'], c['name'], c['old'], c['new'],
                                           '' if c.get('idx_new') is not None else '(騎手指数は旧騎手のまま)')
                   for r, c in q['jk'][:10])
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


# 予想CSVの読込: horselist の再読込・騎手指数表・統計予想の対象日をここでまとめて設定する
_read_racecard_csv_base = read_racecard_csv


def read_racecard_csv(path, require_cols=('場所', '番')):
    global _HL_MISSING_WARNED, _HL_MULTIDATE_WARNED
    df = _read_racecard_csv_base(path, require_cols)
    _HL_MISSING_WARNED = False
    _HL_MULTIDATE_WARNED = False
    try:
        _jockey_day_build(df)
    except Exception:
        pass
    try:
        stat_set_target_date_from_name(Path(str(path)).name)
    except Exception:
        pass
    return df


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


def compute_stat_for_group(group):
    # 1レース分の統計勝率・融合勝率・統計印・参考買い目。結合できなければ None。
    if not (STAT_ENABLE and STAT_DEBA_RACES and _bsm is not None):
        return None
    _m = stat_model()
    if not _m:
        return None
    try:
        _venue = re.sub(r'\s', '', unicodedata.normalize('NFKC', str(group['場所'].iloc[0])))
        _rno = int(unicodedata.normalize('NFKC', str(group['レース'].iloc[0])).replace('R', '').strip())
        _names = {int(b): str(n) for b, n in zip(group['番'], group['馬名']) if pd.notna(b)}
    except Exception:
        return None
    _race = _stat_find_race(_venue, _rno, _names)
    if _race is None:
        return None
    # CSV に無い馬(取消など)は除いてから計算する
    _race2 = _bsm.DebaRace(_race.date, _race.venue, _race.rno, _race.post,
                           [h for h in _race.horses if h['uma'] in _names], _race.path)
    if len(_race2.horses) < 2:
        return None
    _pex = {}
    for b, v in zip(group['番'], group['単勝確率']):
        try:
            if float(v) == float(v):          # ★v136_021 NaN は渡さない
                _pex[int(b)] = float(v)
        except Exception:
            pass
    try:
        _res = _bsm.predict_race(_m, _race2, p_exist=_pex)
    except Exception as _e:
        print(f'  [統計予想] {_venue}{_rno}R の計算に失敗: {_e}')
        return None
    _ps = {b: r['p_stat'] * 100 for b, r in _res.items()}
    _pf = {b: r.get('p_fused', r['p_stat']) * 100 for b, r in _res.items()}
    _order = sorted(_pf, key=lambda b: (-_pf[b], b))
    _marks = {b: STAT_MARKS[i] for i, b in enumerate(_order[:len(STAT_MARKS)])}
    _bets = []
    if len(_order) >= 3:
        _um, _wd = _bsm.pair_probs({b: _pf[b] for b in _order})
        _ax = _order[0]
        for _kind, _tab in (('umaren', _um), ('wide', _wd)):
            for _pt in _order[1:3]:
                _bets.append(dict(kind=_kind, ban1=_ax, ban2=_pt,
                                  prob=round(_tab[frozenset((_ax, _pt))] * 100, 1)))
    return dict(race_key=_race.key, path=_race.path, p_stat=_ps, p_fused=_pf,
                marks=_marks, order=_order, bets=_bets)


def _stat_td(row):
    # 研究用HTMLの馬表に足す3列(統計印・統計%・融合%)
    _mk = str(row.get('統計印', '') or '')
    def _f(v):
        try:
            v = float(v)
            return f'{v:.1f}' if v == v else ''
        except Exception:
            return ''
    _cls = ' class="sthi"' if _mk.startswith('◎') else ''
    _mcls = ' class="sthi uracell"' if STAT_URA_MARK in _mk else _cls
    return f'<td{_mcls}>{_mk}</td><td>{_f(row.get("統計勝率"))}</td><td{_cls}>{_f(row.get("融合勝率"))}</td>'


def _stat_text_lines(analysis):
    """★v136_020: Excel/CSV用。統計裏付けと観察条件(成立分)の説明行。"""
    _out = []
    _ura = analysis.get('stat_ura') or {}
    if _ura:
        _seen = {}
        for _k in ('UR1', 'UR2', 'WD1', 'OU1'):
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
    for _k in ('UR1', 'UR2', 'WD1', 'OU1'):
        _cols += [f'{_k}軸', f'{_k}軸統計%', f'{_k}裏付け']
    _cols += ['本線買い目', '本線裏付け', '次点買い目', '次点裏付け']
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
        _ura = a.get('stat_ura') or {}
        for _k in ('UR1', 'UR2', 'WD1', 'OU1'):
            _u = _ura.get(_k)
            _r[f'{_k}軸'] = _u['ban'] if _u else ''
            _r[f'{_k}軸統計%'] = ('%.1f' % _u['ps']) if _u else ''
            _r[f'{_k}裏付け'] = ('1' if _u['ok'] else '0') if _u else ''
        _main, _main_u = [], []
        for _kind, _rk, _ck in (('wide', 'wide_recs', 'gz_wide_cond'), ('umaren', 'umaren_recs', 'gz_umaren_cond'),
                                ('umatan', 'umatan_recs', 'gz_umatan_cond')):
            _cm = a.get(_ck) or {}
            for _rec in (a.get(_rk) or []):
                _cn = _cm.get((_rec.ban1, _rec.ban2), '')
                if not _cn or gz_tier(_cn) != 'main':
                    continue
                _t = '%s %s %d-%d' % (_cn, _jp[_kind], _rec.ban1, _rec.ban2)
                _main.append(_t)
                if _stat_ura_of(a, _cn):
                    _main_u.append(_t)
        _ji, _ji_u = [], []
        for _jr in (a.get('jiten_recs') or []):
            for _p in _jr.get('partners', []):
                _t = '%s %s %d-%d' % (_jr.get('tag', ''), _jp.get(_jr['kind'], _jr['kind']), _jr['axis_ban'], _p['ban'])
                _ji.append(_t)
                if _stat_ura_of(a, _jr.get('tag')):
                    _ji_u.append(_t)
        _r['本線買い目'] = ' / '.join(_main)
        _r['本線裏付け'] = ' / '.join(_main_u)
        _r['次点買い目'] = ' / '.join(_ji)
        _r['次点裏付け'] = ' / '.join(_ji_u)
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
    _mk = ' '.join('%s%d %s(%.1f%%)' % (_st['marks'][b], b, _h(_nm.get(b, '')), _st['p_fused'][b])
                   for b in _st['order'][:4])
    _jp = {'umaren': '馬連', 'wide': 'ワイド'}
    _bt = ' / '.join('%s %d-%d(%.1f%%)' % (_jp[x['kind']], x['ban1'], x['ban2'], x['prob'])
                     for x in _st['bets'])
    _extra = ''
    _ura = analysis.get('stat_ura') or {}
    if _ura:
        _seen = {}
        for _k in ('UR1', 'UR2', 'WD1', 'OU1'):
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
    return ('<div class="srow"><span class="stag">統計予想（参考・実弾外）</span>'
            '<span class="smk">%s</span><br><span class="sbt">参考買い目 %s</span>%s</div>' % (_mk, _bt, _extra))
REF_N_PARTNERS = 3             # 参考買い目の点数(=○▲△の頭数)
REF_VALUE_LAMBDA = 4.0         # 妙味の重み λ(★v136_015: 1.0→4.0 回収率重視。1.0=的中重視)
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
        return '投資レース', '軸の信頼度が特に高いレース(推定3着内率88%以上)。回収率100%超の保証はない'
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


def compute_axis_grade(複合スコア, gap, 複勝確率, 単勝確率, 単勝期待値, 紐馬指数):
    """最適化式で軸馬区分('S級軸'..'D級軸')を返す。"""
    _v = {'複合スコア': 複合スコア, 'gap': gap, '複勝確率': 複勝確率, '単勝確率': 単勝確率,
          '単勝期待値': min(単勝期待値, 160.0), '紐馬指数': 紐馬指数}
    score = 0.0
    for _f in _AGF_FEATS:
        _sd = _AGF_STD.get(_f) or 1.0
        score += _AGF_W.get(_f, 0.0) * ((_v.get(_f, 0.0) - _AGF_MEAN.get(_f, 0.0)) / _sd)
    for _g, _c in _AGF_CUTOFFS:
        if _c is None or score >= _c:
            return _g
    return _AGF_CUTOFFS[-1][0]

def compute_axis_grade_letter(複合スコア, gap, 複勝確率, 単勝確率, 単勝期待値, 紐馬指数):
    """既存コードが前提とする英字グレード('S'..'D')を返す。"""
    g = compute_axis_grade(複合スコア, gap, 複勝確率, 単勝確率, 単勝期待値, 紐馬指数)
    return (g.replace('級軸', '').strip() or 'D')


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


def get_unfold_score(unfold):
    unfold = str(unfold)
    if '逃' in unfold: return 100
    elif '先' in unfold: return 90
    elif '差' in unfold: return 80
    elif '追' in unfold: return 60
    else: return 70

def get_dynamic_unfold_score(unfold, distance):
    base = get_unfold_score(unfold)
    dist_cat = 'short' if distance <= 1400 else 'middle' if distance <= 1800 else 'long'
    bonus = {
        'short':  {'逃': +15, '先': +10, '差': -5,  '追': -10},
        'middle': {'逃': +5,  '先': +8,  '差': +5,  '追': 0},
        'long':   {'逃': -10, '先': -5,  '差': +10, '追': +15}
    }.get(dist_cat, {})
    for key, val in bonus.items():
        if key in str(unfold):
            base += val
    return max(50, min(120, base))


def race_zscore(series):
    s = pd.Series(series).reset_index(drop=True)
    mean = s.mean()
    std = s.std(ddof=0)
    return (s - mean) / std if std != 0 else pd.Series(0, index=s.index)

# ══════════════════════════════════════════════════════════════════
# ★v135_005: 表示専用『補正騎手指数』(公開用の参考値)
#   目的: 買い目条件に使う生の 騎手指数 は部外秘のため表に出せない。
#         そこで単勝オッズで補正し偏差値化した値を参考として公開する。
#
#   ★この値は買い目条件には一切使わない(条件は生の騎手指数のまま)。
#     算出も table 表示用の列としてのみ行う。
#
#   mode='resid' (既定):
#     レース内で 騎手指数 を log10(単勝オッズ) に単回帰し、その残差を偏差値化。
#     = 「オッズ(市場の人気)で説明できない分の騎手評価」。
#     人気馬に乗っているから指数が高いだけの騎手は 50 付近に寄る。
#   mode='ratio':
#     騎手指数 × log10(単勝オッズ+1) を偏差値化(妙味方向に素直に効く単純版)。
#
#   偏差値: 50 + 10 × (x - 平均) / 標準偏差。母集団はそのレースの出走馬のみ。
#   出走3頭未満 / 標準偏差0 / 値が欠損 の場合は 50.0 を返す。
#
#   ※注意(情報保護の限界): 単勝オッズは公開情報なので、この偏差値と
#     オッズが揃えばレース内の生騎手指数はアフィン変換の範囲まで復元できる。
#     絶対水準と縮尺は隠れるが順位と相対間隔は漏れる。完全な秘匿ではない。
# ══════════════════════════════════════════════════════════════════
KISHU_DEV_COL   = '騎手補正'      # 内部列名(table_df の列名)
KISHU_DEV_LABEL = '騎手補正'      # 表示ラベル(HTML/Excel/GUIの見出し)


def calc_kishu_dev(group, mode=None):
    """表示専用の補正騎手指数(偏差値)を返す。index は group と同じ。"""
    if mode is None:
        mode = str(getattr(Config, 'KISHU_DEV_MODE', 'resid'))
    n = len(group)
    fallback = pd.Series(50.0, index=group.index)
    if n < 3 or '騎手指数' not in group.columns:
        return fallback
    x = pd.to_numeric(group['騎手指数'], errors='coerce')
    _oc = '単勝オッズ' if '単勝オッズ' in group.columns else 'adjusted_odds'
    if _oc not in group.columns:
        return fallback
    o = pd.to_numeric(group[_oc], errors='coerce')
    # オッズは 1.0倍未満・欠損を除外対象にする
    valid = x.notna() & o.notna() & (o >= 1.0)
    if int(valid.sum()) < 3:
        return fallback
    lo = np.log10(o.where(valid).clip(lower=1.0))
    if mode == 'ratio':
        raw = x.where(valid) * np.log10(o.where(valid).clip(lower=1.0) + 1.0)
    else:
        # レース内で 騎手指数 を log10(オッズ) に単回帰し残差を取る
        xv, ov = x[valid].astype(float), lo[valid].astype(float)
        ov_var = float(((ov - ov.mean()) ** 2).sum())
        if ov_var <= 0:
            raw = xv.reindex(group.index)
        else:
            beta = float(((ov - ov.mean()) * (xv - xv.mean())).sum()) / ov_var
            alpha = float(xv.mean() - beta * ov.mean())
            raw = (x.where(valid) - (alpha + beta * lo))
    rv = pd.to_numeric(raw, errors='coerce')
    sd = rv.std(ddof=0)
    if pd.isna(sd) or sd == 0:
        return fallback
    dev = 50.0 + 10.0 * (rv - rv.mean()) / sd
    # 偏差値は 20〜80 に収めて表示(外れ値でレイアウトが崩れるのを防ぐ)
    return dev.clip(lower=20.0, upper=80.0).fillna(50.0).round(1)


def race_percentile(series):
    s = pd.Series(series).reset_index(drop=True)
    return s.rank(pct=True) * 100

def calculate_ics(group):
    group = group.reset_index(drop=True)
    ability_raw = (0.40 * group['単指数'] + 0.45 * group['複指数'] + 0.15 * group['騎手指数'])
    p_ability = race_percentile(ability_raw)
    p_jockey = race_percentile(group['騎手指数'])
    p_unfold = race_percentile(group['unfold_score'])
    tactical = 0.50 * p_jockey + 0.50 * p_unfold
    edge_raw = group['単指数'] / group['adjusted_odds'].clip(lower=1.0)
    p_edge = race_percentile(edge_raw)
    pos_raw = 18 - group['番']
    p_pos = race_percentile(pos_raw)
    ics_score = (0.40 * p_ability + 0.25 * tactical + 0.25 * p_edge + 0.10 * p_pos)
    return round(ics_score.clip(0, 100), 1)


def get_joint_synergy_bonus(single_idx, fuku_idx):
    if pd.isna(single_idx) or pd.isna(fuku_idx): return 0.0
    s = float(single_idx)
    f = float(fuku_idx)
    bonus = 0.0
    if f >= 25:
        if s >= 94: bonus = 7.5
        elif s >= 85: bonus = 6.0
        elif s >= 76: bonus = 4.5
        else: bonus = 3.0
    elif f >= 23:
        if s >= 94: bonus = 5.5
        elif s >= 85: bonus = 4.5
        elif s >= 76: bonus = 3.5
        else: bonus = 2.5
    elif f >= 20:
        if s >= 85: bonus = 2.0
        elif s >= 70: bonus = 1.5
    if s >= 100 and f >= 22: bonus += 1.5
    elif s >= 95 and f >= 21: bonus += 1.0
    return round(min(bonus, 9.0), 2)


def wide_odds_from_place(place_a, place_b, takeout=Config.WIDE_TAKEOUT):
    """★ v131.1: ワイド予測オッズを『両馬の複勝確率』から算出する正しい式。
    ワイドは2頭がともに3着内に入れば的中する券なので、的中確率は各馬の複勝率で決まる。
    旧 estimate_realistic_odds(単勝オッズの和ベース)は同時複勝の広さを反映せず、
    複勝確率ベースの理論値の2〜4倍に系統的に過大化していたため置換する。
      place_a, place_b : 各馬の複勝確率(%). エンジン算出の '複勝確率' 列を渡す。
      戻り値           : 推定ワイドオッズ(払戻率 takeout を織り込んだ期待オッズ)。
    近似: P(A3着内 かつ B3着内) ≈ (pA/100)*(pB/100)。強い2頭の負相関で実際はやや広く
    出るため、係数 corr(<1)で同時確率を軽く割り引き、控除後オッズ = takeout / P とする。
    """
    pa = max(0.0, min(100.0, float(place_a))) / 100.0
    pb = max(0.0, min(100.0, float(place_b))) / 100.0
    joint = pa * pb * Config.WIDE_JOINT_CORR
    joint = min(0.985, max(1e-4, joint))
    return round(max(1.1, takeout / joint), 1)


def estimate_realistic_odds(axis_odds, target_odds, target_pop, arare_index, dist, head_count, bet_type='wide'):
    base_wide = (axis_odds + target_odds) / Config.WIDE_ODDS_DIVISOR
    pop_diff_bonus = 1.32 if target_pop >= 6 else 1.18 if target_pop >= 4 else 1.08 if target_pop >= 2 else 1.0
    arare_factor = 1.22 if arare_index >= 65 else 1.08 if arare_index >= 45 else 0.92
    if bet_type == 'umatan': arare_factor = 1.35 if arare_index >= 65 else arare_factor
    dist_factor = 1.12 if dist <= 1400 else 0.85 if dist >= 1800 else 1.0
    head_factor = 1.0 + (18 - head_count) * 0.015 if head_count < 12 else 1.0
    est_wide = round(base_wide * pop_diff_bonus * arare_factor * dist_factor * head_factor, 1)
    if bet_type == 'wide': return max(1.1, est_wide)
    elif bet_type == 'umaren': return round(est_wide * Config.UMAREN_RATIO, 1)
    elif bet_type == 'umatan': return round(est_wide * Config.UMAREN_RATIO * Config.UMATAN_RATIO, 1)
    return est_wide


# 穴馬軸の人気下限(これ以上の人気薄を「穴」とみなす)
_HOLE_MIN_POP = 6

def detect_hole_candidates(group, axis_set):
    """穴馬軸推奨(v091): N数の多い『有力な単勝条件』に該当した人気薄馬を抽出する。
    これらを複勝馬として表示し、2連系(馬連/馬単/ワイド)の軸ロジックを維持する。
    条件: 人気薄(補正後人気>=_HOLE_MIN_POP) かつ
      ・有力単勝条件(単YES, 回収率>=_AXIS_TAN_MIN_ROI=105, N>=_AXIS_MIN_N)に該当、または
      ・複勝表示対象条件(複YES: N数の多いB/C/D)に該当。
    採用条件は内蔵の検証済みテーブル(_AXIS_COND_FALLBACK)で解決。"""
    res = []
    for i in group.index:
        if i in axis_set:
            continue
        pop = safe_pop(group.loc[i, 'adjusted_popularity'])
        if not (pop >= _HOLE_MIN_POP):   # NaN(欠損)も除外
            continue
        at = _axis_attr(group, i)
        q1, r1, _n1, _c1 = eval_axis_quality(at, '単')
        q2, _r2, _n2, _c2 = eval_axis_quality(at, '複')
        if (q1 and r1 >= _AXIS_TAN_MIN_ROI) or q2:
            res.append(i)
    return res


# =====================================================
# ★ v083: 紐馬「好条件」エンジン（optimal_himo_conditions.csv 由来）
#   実データ・バックテストで回収率100%超(YES)となった紐条件を読み込み、
#   各紐馬が券種別(馬連/馬単/ワイド)の好条件を満たすかを判定する。
#     - 該当券種で100超(YES)条件を1つでも満たす → その券種の「好条件紐」
#     - 強条件(高回収率)を満たす → 「有力紐」。最大4点買いの可否に使う。
#   CSVが見つからない場合は内蔵フォールバック条件で動作する。
#   評価値はエンジン算出値: 複勝率=複勝確率, EV=単勝期待値, オッズ=adjusted_odds,
#   複指数=複指数, 人気=adjusted_popularity, 馬指数=紐馬指数。
# =====================================================
import re as _re


# ★ 紐条件テーブルの構築状況（算出内訳表示用）


# =====================================================
# ★ v085: 軸馬「単勝/複勝」買い判定エンジン（optimal_axis_conditions.csv 由来）
#   実データ・バックテストで単勝/複勝の回収率100%超(YES)となった軸条件を読み込み、
#   NCSで選んだ馬が「実際に買ってプラスの条件」に該当するかをゲート判定する。
#     - 単勝100超(YES) かつ N>=_AXIS_MIN_N → 単勝の買い対象
#     - 複勝100超(YES) かつ N>=_AXIS_MIN_N → 複勝の買い対象
#   NCS軸が該当=本命単/複、軸以外の該当馬=穴単/穴複として拾う。
#   評価値はエンジン算出値: 単指数, 複指数, オッズ=adjusted_odds,
#   人気=adjusted_popularity, 馬番=番, 脚質=展開。
#   CSVが無い場合は内蔵フォールバック(主要YES条件)で動作する。
# =====================================================
_AXIS_MIN_N = 40          # v092: 複勝の安定条件(N=40-64)も拾えるよう40に調整
_AXIS_TAN_MIN_ROI = 105.0 # 単勝は条件回収率>=105%を必須(安定条件のみ採用)

# 内蔵フォールバック: (条件, N, 単勝回収, 複勝回収, 単YES, 複YES)
_AXIS_COND_FALLBACK = [
    # v092(Opus4.8 全面改訂): optimal_axis_conditions.csv の実配当バックテストから
    #   ・回収率 >= 105% ・N(サンプル数) >= 150(複勝のみ >=40) の安定条件のみ採用。
    #   共通原理 = 指数(単/複)が裏付ける過小評価馬。脚質追×複指数23-25帯が単勝の主力。
    #   (条件, N, 単勝回収率%, 複勝回収率%, 単YES, 複YES)
    # ── 単勝(YES単): 回収105%超・N>=150。安定(高頻度)条件を中心に高回収条件を併用。
    ('脚質追 + 複指数23-25',                       1106, 111.4,  99.4, 1, 0),  # 主力・最大N(勝率15.1%)
    ('脚質追 + 複指数23-25 + 人気4+',               781, 115.8,  90.0, 1, 0),
    ('馬番7+ + 脚質追 + 複指数23-25',               515, 134.0,  99.4, 1, 0),  # 勝率17.3%
    ('脚質追 + 複指数23-25 + オッズ10-30',          432, 136.6,  89.3, 1, 0),
    ('複指数20+ + オッズ30+ + 人気1-5',             250, 181.0,  78.8, 1, 0),  # 高回収・低頻度
    ('馬番1-6 + オッズ30+ + 人気1-5',               271, 163.9,  87.2, 1, 0),
    ('単指数80+ + オッズ<10 + 人気4+',              238, 111.6,  90.0, 1, 0),  # 勝率19.8%
    ('複指数20+ + オッズ<10 + 人気6+',              208, 121.4,  90.0, 1, 0),
    ('馬番7+ + オッズ<10 + 人気6+',                 161, 118.6,  90.0, 1, 0),
    ('単指数60+ + オッズ<10 + 人気6+',              159, 132.8,  96.1, 1, 0),
    # ── 複勝(YES複): 回収105%超。母数は単勝より小さいが安定条件のみ。
    ('脚質逃 + 複指数23-25 + 人気6+',                63,  90.0, 122.5, 0, 1),  # 複勝率25.4%
    ('単指数70+ + オッズ<10 + 人気6+',               64,  90.0, 108.3, 0, 1),  # 複勝率43.8%
    ('複指数23-25 + オッズ<10 + 人気6+',             49,  90.0, 109.2, 0, 1),  # 複勝率42.9%
    ('馬番1-3 + 脚質追 + 単指数80+',                 40, 114.1, 110.3, 0, 1),  # 複勝率70%
    ('馬番1-4 + 脚質追 + 単指数80+',                 50, 114.1, 106.8, 0, 1),  # 複勝率70%
]

def _parse_axis_token(tok):
    """軸条件の1トークンを述語(attr->bool)に変換。脚質と数値メトリクスに対応。"""
    tok = tok.strip()
    m = _re.match(r'脚質(差|追|逃|先)$', tok)
    if m:
        ch = m.group(1)
        return lambda a: ch in str(a.get('leg', ''))
    m = _re.match(r'(単指数|複指数|複合順位|オッズ|人気|馬番)(.+)$', tok)
    if not m:
        return None
    metric, rest = m.group(1), m.group(2).strip()
    key = {'単指数': 'tanI', '複指数': 'fukushisu', 'オッズ': 'odds',
           '人気': 'pop', '馬番': 'ban', '複合順位': 'csr'}[metric]
    rng = _re.match(r'^(\d+)\s*-\s*(\d+)$', rest)
    if rng:
        lo, hi = float(rng.group(1)), float(rng.group(2))
        return lambda a: lo <= a[key] <= hi
    lt = _re.match(r'^<\s*(\d+)$', rest)
    if lt:
        v = float(lt.group(1)); return lambda a: a[key] < v
    pl = _re.match(r'^(\d+)\s*\+$', rest)
    if pl:
        v = float(pl.group(1)); return lambda a: a[key] >= v
    eq = _re.match(r'^(\d+)$', rest)
    if eq:
        v = float(eq.group(1)); return lambda a: a[key] == v
    return None

def _parse_axis_condition(cond_str):
    preds = []
    for tok in _re.split(r'\s+\+\s+', str(cond_str).strip()):
        if not tok.strip():
            continue
        p = _parse_axis_token(tok)
        if p is None:
            return None
        preds.append(p)
    if not preds:
        return None
    return lambda a: all(p(a) for p in preds)

_AXIS_COND_CACHE = None
# ★ optimal_axis_conditions.csv の読込み状況（算出内訳表示用）
_AXIS_COND_STATUS = {'loaded': False, 'path': None}
def _load_axis_conditions():
    global _AXIS_COND_CACHE
    if _AXIS_COND_CACHE is not None:
        return _AXIS_COND_CACHE
    import os as _os
    rows = []
    cand_paths = []
    envp = _os.environ.get('AXIS_COND_PATH')
    # v091: 既定では内蔵の「検証済み 有力単複条件」を正とする。
    #   旧 optimal_axis_conditions.csv は小N過剰適合のため、AXIS_COND_PATH を
    #   明示指定した場合のみ読み込む(自動探索は無効化)。
    if envp:
        cand_paths.append(Path(envp))
    loaded = False
    for p in cand_paths:
        try:
            if p and Path(p).exists():
                with open(p, 'rb') as _fh:
                    _raw = _fh.read().decode('utf-8-sig', errors='ignore')
                # v097: 先頭の # コメント行(軸馬・紐馬の選び方の注記)はスキップ
                _ls = [ln for ln in _raw.splitlines() if not ln.lstrip().startswith('#')]
                _nc = _ls[0].count(',') if _ls else 0
                _good = [ln for ln in _ls if ln.count(',') == _nc]
                import io as _io2
                _df = pd.read_csv(_io2.StringIO('\n'.join(_good)))
                for _, r in _df.iterrows():
                    pred = _parse_axis_condition(r['条件'])
                    if pred is None:
                        continue
                    rows.append(dict(
                        cond=str(r['条件']), n=int(r['N']), pred=pred,
                        roi={'単': float(r['単勝回収率']), '複': float(r['複勝回収率'])},
                        yes={'単': str(r['単勝100超']).strip().upper() == 'YES',
                             '複': str(r['複勝100超']).strip().upper() == 'YES'}))
                if rows:
                    loaded = True
                    _AXIS_COND_STATUS['loaded'] = True
                    _AXIS_COND_STATUS['path'] = str(p)
                    break
        except Exception:
            continue
    if not loaded:
        for (cond, n, sr, fr, sy, fy) in _AXIS_COND_FALLBACK:
            pred = _parse_axis_condition(cond)
            if pred is None:
                continue
            rows.append(dict(cond=cond, n=n, pred=pred,
                             roi={'単': sr, '複': fr},
                             yes={'単': bool(sy), '複': bool(fy)}))
    _AXIS_COND_CACHE = rows
    return rows

def _compound_rank(group, i):
    """複合スコア = 単指数*0.7 + BADO指数*0.3 のレース内 降順順位(1位=最良)。
    分析(analysis_horse_detail.csv)の複合スコア順位を完全再現する定義。"""
    try:
        cs = group['単指数'].astype(float) * 0.7 + group['BADO指数'].astype(float) * 0.3
        return int(cs.rank(ascending=False, method='min').loc[i])
    except Exception:
        return 99

def _axis_attr(group, i):
    # v091b: 採用条件は『素の市場オッズ・人気』で検証したものなので、
    #   条件判定にも補正後ではなく元の単勝オッズ/単勝人気を用いて再現性を確保する。
    # ★欠損(0埋め)は NaN 扱い。0 のままだと『オッズ<10』『人気1-5』等の
    #   条件を素通りしてしまうため(NaN なら全比較が False になり安全側)。
    _raw_od = (safe_odds(group.loc[i, '単勝オッズ']) if '単勝オッズ' in group.columns
               else safe_odds(group.loc[i, 'adjusted_odds']))
    _raw_pop = (safe_pop(group.loc[i, '単勝人気']) if '単勝人気' in group.columns
                else safe_pop(group.loc[i, 'adjusted_popularity']))
    return dict(
        tanI=float(group.loc[i, '単指数']), fukushisu=float(group.loc[i, '複指数']),
        odds=_raw_od, pop=_raw_pop,   # ※どちらも欠損時は NaN(全条件が False)
        ban=int(group.loc[i, '番']), leg=str(group.loc[i, '展開']),
        csr=_compound_rank(group, i))

def eval_axis_quality(attr, kind):
    """kind: '単' or '複'。該当(YES & N>=下限)するか・最良回収率・N・条件名を返す。"""
    conds = _load_axis_conditions()
    best_roi = 0.0; best = None; n_best = 0; q = False
    for c in conds:
        if not c['yes'].get(kind, False):
            continue
        if c['n'] < _AXIS_MIN_N:
            continue
        try:
            ok = c['pred'](attr)
        except Exception:
            ok = False
        if not ok:
            continue
        q = True
        roi = c['roi'].get(kind, 0.0)
        if roi > best_roi:
            best_roi = roi; best = c['cond']; n_best = c['n']
    return q, best_roi, n_best, best


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
# ★ 参考強制ゲート(買い目は出すが備考を『参考』にする)。
#   旧ゲートは複合スコア軸時代の実配当検証に基づくもので、現行の
#   確率ベース+採用条件には対応しないため空集合で無効化してある。
#   新ロジックでの検証結果が出たら必要に応じて再設定する。
NSL_REFERENCE_ONLY_VENUES = set()   # 参考強制する開催地: なし
NSL_TIER_REFERENCE_ONLY = set()     # 参考強制する信頼度区分: なし


# ════════════════════════════════════════════════════════════════
# ★ v134: トリガミ覚悟の「安定運用」単勝/複勝 買い目
# ────────────────────────────────────────────────────────────────
#   一連の検証(bootstrap CI下限, train/valid分離)で見えた
#   「勝てはしないが的中率が高く下振れが浅い」領域を買い目化する。
#
#   軸馬(=単勝確率1位)が下記4条件を満たすとき、開催地に応じて
#   単勝/複勝を『安定』として表示する。
#     ① 単勝確率(win_prob)   >= S4_WIN_PROB_MIN (既定50%)
#     ② 複勝確率(place_prob) >= S4_PLACE_PROB_MIN(既定70%  ※80は絞りすぎ)
#     ③ 補正前単勝人気        <= S4_RAW_POP_MAX  (既定3)
#     (④ 単勝確率1位 は _tan_target_idx により保証)
#
#   開催地ゲート(検証の複勝/単勝CI下限・回収率より):
#     複勝安定 = 盛岡/園田/金沢/笠松/名古屋/佐賀 (門別/大井/水沢は除外)
#     単勝安定 = 盛岡/名古屋/園田/金沢/笠松       (高知/門別は除外)
#   ゲート外の開催地でも条件を満たせば『参考(安定外)』として出す(消さない)。
#   値は yuryoku_conditions.json の "stable_ops" キーで上書き可能。
# ════════════════════════════════════════════════════════════════
S4_WIN_PROB_MIN   = 50.0
S4_PLACE_PROB_MIN = 70.0
S4_RAW_POP_MAX    = 3
# 複勝で回収率が相対的に高い(CI下限が高い)開催地
# ★v136_018: 単勝・複勝【安定】を本線から外す(実戦条件シートに含まれないため)。True で復活。
STABLE_BET_ENABLE = False
STABLE_FUKU_VENUES = {'盛岡', '園田', '金沢', '笠松', '名古屋', '佐賀'}
# 単勝で回収率が相対的に高い開催地
STABLE_TAN_VENUES  = {'盛岡', '名古屋', '園田', '金沢', '笠松'}


def _load_stable_ops_config():
    """yuryoku_conditions.json の 'stable_ops' があれば安定運用パラメータを上書き。"""
    global S4_WIN_PROB_MIN, S4_PLACE_PROB_MIN, S4_RAW_POP_MAX
    global STABLE_FUKU_VENUES, STABLE_TAN_VENUES
    try:
        here = _os.path.dirname(_os.path.abspath(__file__))
        for d in (here, '.', _os.environ.get('PROB_MODEL_DIR') or ''):
            if not d:
                continue
            p = _os.path.join(d, 'yuryoku_conditions.json')
            if _os.path.exists(p):
                with open(p, encoding='utf-8') as f:
                    cfg = _json.load(f)
                so = cfg.get('stable_ops')
                if isinstance(so, dict):
                    S4_WIN_PROB_MIN = float(so.get('win_prob_min', S4_WIN_PROB_MIN))
                    S4_PLACE_PROB_MIN = float(so.get('place_prob_min', S4_PLACE_PROB_MIN))
                    S4_RAW_POP_MAX = int(so.get('raw_pop_max', S4_RAW_POP_MAX))
                    if so.get('fuku_venues'):
                        STABLE_FUKU_VENUES = set(so['fuku_venues'])
                    if so.get('tan_venues'):
                        STABLE_TAN_VENUES = set(so['tan_venues'])
                return
    except Exception:
        pass


_load_stable_ops_config()


def _stable_axis_ok(win_prob, place_prob, raw_pop):
    """軸馬が安定運用の4条件(①②③＋1位)を満たすか。1位判定は呼び出し側。"""
    try:
        if win_prob is None or place_prob is None or raw_pop is None:
            return False
        # ★raw_pop<=0 は「人気欠損」であり1番人気ではない → 不成立にする
        return (float(win_prob) >= S4_WIN_PROB_MIN
                and float(place_prob) >= S4_PLACE_PROB_MIN
                and 1 <= int(raw_pop) <= S4_RAW_POP_MAX)
    except Exception:
        return False


# ══════════════════════════════════════════════════════════════════
# ★v135_003: 買い目条件レジストリ(全面刷新)
#   prob_model_conditions_with_axis.json の robust条件(学習ROI>=100 かつ 検証ROI>=100 かつ 検証N>=30)を、
#   ①検証的中数 ②学習と検証の的中率一致 ③平均配当が控除率に負けない帯
#   の3点で選抜した 馬単6条件 / 馬連3条件 に置き換える。
#
#   ※単勝/複勝(安定運用)は従来どおり別系統(_stable_axis_ok)で維持。
#
#   軸ルール(JSONと同一): 軸atomを満たす馬のうち 単勝確率_生 が最大の馬。
#   atom解釈: win_prob=単勝確率_生 / place_prob=複勝確率_生 / tan_idx=単指数
#             kishu_idx=騎手指数 / odds=補正前単勝オッズ / ev=EV_生
#
#   stat: p =1点あたり的中率(学習+検証プール) / pt=点ROI(%) /
#         lo=ROI下限の近似(%) / hit=的中数(プール)
#   ★lo は「的中率のWilson下限 × 平均配当」で算出した近似値であり、
#     払戻額の分散を無視しているため実際のレース単位ブロックブートストラップ
#     CI下限より楽観的に出る。min(プール下限, 検証下限)を採用して安全側に寄せて
#     いるが、いずれの条件も統計的に利益が確認されたわけではない点は旧条件と同じ。
# ══════════════════════════════════════════════════════════════════
# ─────────────────────────────────────────────────────────────
# ★F4条件の係数（v16 [8] c/b0感度検査の結果、固定値に確定）
#
#   旧: f4_params.json を fit_f4_params.py で毎回推定して読み込む
#   新: 下記の定数を直接使う。JSONもフィッタも不要。
#
#   固定値化の根拠 (himo_intersect_v13.json / cb0_sweep):
#     ・c を 1.442〜5.0 の全域で「正直なCI下限>=100% かつ Aに優位」を満たす
#     ・b0 を -4.0〜-2.0 の全域で同様に満たす
#     ・固定値(正直CI下限104.0%)が月次推定(103.0%)をわずかに上回る
#       → 毎月の再推定に価値がないことが確認された
#     ・slope は [3] で 0.30〜0.90 の全域で成立済み
#   両端(c=0.4 / b0=-1.2)で崩れるのは仕様どおり:
#     c小 → sigmoidが平坦化しオッズ最大馬が選ばれS2との積が空に近づく
#     c大 → 複勝確率が支配的になりS2と同一集合に近づく
#
#   ★この定数を「良さそうな値」に手で動かさないこと。合格域の中で
#     微調整すると選択バイアスを持ち込み、検証済みという性質が失われる。
#     変更したい場合は Himo_recent_v16.py の [8] を回し直してから。
#
#   ※fit_f4_params.py と f4_params.json は不要になったため削除してよい。
#     f4_params.json が残っていても読まれない(意図的に無視する)。
# ─────────────────────────────────────────────────────────────
import math as _math

F4_C = 1.44194956741415
F4_B0 = -2.6862502942403292
F4_SLOPE = 0.6164341331784238
_F4_PARAMS = dict(c=F4_C, b0=F4_B0, slope=F4_SLOPE)
_F4_ERR_WARNED = False


def _load_f4_params():
    """★固定係数を返す。外部ファイルは読まない(v135_018で廃止)。

    以前は f4_params.json を探し、無ければ None を返して F4 を無効化
    していた。ファイル欠損で本線条件が黙って落ちる事故要因だったが、
    係数が固定値でよいと確認されたためその経路ごと不要になった。
    呼び出し側の `if not _P: return []` は残してある(構造は不変)。
    """
    return _F4_PARAMS


GZ_COND_DEFS = {
    # ══ 本線枠 (tier='main') ══════════════════════════════════════
    #   採用基準: 軸atomが win_prob>=40/50 に限定され、検証ROI>=110% かつ
    #   1点あたり的中率が高いもの。旧U2/U7はU3とほぼ同一集合のため統合し、
    #   旧R1(馬連・軸ev1.5+)は軸的中率12.8%で他条件と土俵が違うため削除。
    #
    #   ★v135_015 で削除した条件: U2 / U7 / R1 / A20 / A25 / A30 / A40 / A50
    #     - U2,U7: U3と軸・相手の大半が重複(軸wp50+相手複40)。3条件分の
    #       点数を投じても外れるときは同時に外れる相関リスクだけが増える。
    #     - R1: 軸atomが ev1.5+ で軸自体が飛ぶ設計。被覆率集計で
    #       発火180R中 軸的中23R(12.8%)しかなく、本線枠の基準を満たさない。
    #     - A20/A25/A30: 軸が wp>=20〜30 と緩く、本線枠が発火しないレースで
    #       単独発火していた(紐カバーではなく別物の穴狙い)。増分ROI 66.8%。
    #     - A40/A50: カバー枠としては C1-C3/E1-E3 が全項目で上回るため置換。
    # ══════════════════════════════════════════════════════════
    # ★v135_019 の変更 (2026/08)  検証: --himo-coverage valid 2,372R
    #
    #   (1) F4/S2 の軸を win_prob>=40 → >=50 に変更
    #   (2) U3 / C3 / E2 を削除
    #
    #   実測比較（本線枠のみ / 全枠）:
    #     v135_018        106.4% (CI下限79.2) / 101.2% (CI下限79.8)  投資251,500円
    #     (1)のみ         122.8% (CI下限86.6) / 108.6% (CI下限82.9)  投資215,200円
    #     (1)+(2)         122.3% (CI下限85.9) / 115.1% (CI下限86.6)  投資198,800円
    #   → 投資を21%減らしてROIが14pt改善。リスクを下げて成績が上がった。
    #
    #   ★軸wp50化の根拠: 相手条件を固定して軸だけ wp40→50 に変えた
    #     278組の比較で、検証ROI合成 84.5%→94.9%、278組中233組(84%)で改善。
    #     複勝確率・順位ベースの相手(S2/F4)は軸と情報源が同じため、軸が
    #     緩いと共倒れする。単指数×騎手指数(U8/M1/U1)は軸と独立な情報
    #     なので wp40 のままでよい(締めると配当妙味が消えて悪化する)。
    #     ★F4 ⊆ S2 の包含関係は両者の軸が同一であることに依存する。
    #       片方だけ変えると包含が壊れる。必ずセットで扱うこと。
    #
    #   ★削除の根拠:
    #     E2: 113点で的中2本、ROI 19.0%、CI上限56.6%。上限が100%を大きく
    #         割っており「高配当の紐抜けを埋める」設計意図を果たせていない。
    #         累積表でE2を足した時点のみ 115.1%→108.6% と明確に低下。
    #     C3: 実質29点(取り分率3.9%)。U1と91.9%重複し独自の買い目がない。
    #         ROI 51.8% / CI上限94.8% で単体でも不合格。
    #     U3: 実質11点(取り分率6.2%)。S2/C1/C2と93.8%重複。あっても
    #         無くても買い目がほぼ変わらない。
    #     ※削除しても紐カバー枠の増分ROIは 78.6%→95.1% に改善する。
    #
    #   ★lo は「レース単位の独立ブートストラップ」によるCI下限であり、
    #     月ごとの環境差(tau)を織り込んでいないため楽観的に出る。
    #     階層ブートストラップで補正すると更に5〜8pt下がる。
    #     どの条件も CI下限>=100% には達していない。実弾は少額に留めること。
    # ══════════════════════════════════════════════════════════
    'F4': dict(kind='umatan', tier='main', axis='win_prob>=50', order=-1,
               label='F4 馬単 軸wp50+×複勝率上位2頭∩期待値上位2頭',
               short='軸wp50+ × 複勝上位2 ∩ EV上位2 (1-2点)',
               bg='#9fd8a8', fg='#14532d', fill='9FD8A8',
               p=0.1784, pt=124.0, lo=80.7, hit=43,
               troi=0.0, vroi=124.0, vn=241, vhit=43, avg=695),
    'S2': dict(kind='umatan', tier='main', axis='win_prob>=50', order=0,
               label='S2 馬単 軸wp50+×複勝率上位2頭',
               short='軸wp50+ × 複勝確率上位2頭(2点)',
               bg='#ffcf70', fg='#6b3d00', fill='FFCF70',
               p=0.2424, pt=93.0, lo=68.1, hit=56,
               troi=0.0, vroi=93.0, vn=231, vhit=56, avg=384),
    # ★U3削除(v135_019): 実質11点・取り分率6.2%。S2/C1/C2と93.8%重複。
    # ★v135_020: U8/M1/U1 を U1 単独に統合。
    #   3条件は単指数の閾値(70/60/50)だけが違う入れ子(U8⊆M1⊆U1)で、
    #   買い目は order順の先着優先で重複排除され、合計は U1(単50騎50)
    #   単独と1点も変わらなかった(--cond-overlap で実測: U8実質212点/
    #   M1実質343点/U1実質431点=合計986点 は tan>=50&騎50 の単独点数と一致)。
    #   賭け金は全条件均等単価(Kelly撤去済み)のため、統合で収支は不変。
    #   帯ごとに賭け金を変える設計に戻す場合のみ3条件へ再分割すること。
    # ══════════════════════════════════════════════════════════
    # ★v135_023: 本線枠を A1 / C1 の2条件に差し替え。
    #
    #   A1 = 軸[tan_idx>=90 & odds<=1.5] × 複勝確率上位2頭 (馬単2点)
    #     ・軸の期間耐性がこれまでで最良。単勝/複勝の回収率が
    #       学習と検証で1〜2ptしか乖離しない(win_prob軸は9〜16pt乖離)。
    #       軸単体の複勝: 学習93.5% / 検証99.8%(的中96.7%)。
    #     ・馬単マージンが学習+2.50 / 検証+7.71 と両期間プラス。
    #       255通りの相手探索で符号が揃った数少ない条件。
    #     ・★該当は検証61R(2.6%)と稀少。CI下限91.7で100%は未達。
    #
    #   W1 = 軸[win_prob>=40] × 単指数50&騎手指数50 × 出走11頭以上
    #     ・旧U1に頭数ゲートを追加したもの。頭数別で 11-12頭122.6% /
    #       9-10頭114.2% / ≤8頭67.1% と明確な差(検証)。
    #     ・ただし的中率は各帯25〜26%でほぼ同じ。差は平均配当
    #       (960/873/723円)による。頭数自体がエッジではなく、
    #       多頭数のほうが配当が高いという構造。
    #     ・★軸wp40は単勝マージンが学習+6.08→検証-9.90と16pt乖離する。
    #       この条件だけ両期間122.6%なのは軸の劣化を相手と頭数が
    #       偶然打ち消した形で、構造的な説明がついていない。
    #       CI下限75.9。A1より該当は多い(8.5%)がリスクは高い。
    #
    #   ★どちらもCI下限<100%。実弾は損失許容の範囲に留めること。
    # ══════════════════════════════════════════════════════════
    'A1': dict(kind='umatan', tier='main', axis='tan90_odds15', order=0,
               label='A1 馬単 軸[単指数90+&オッズ1.5以下]×複勝確率上位2頭',
               short='軸 単指数90+ & オッズ1.5以下 × 複勝上位2 (2点)',
               bg='#ffcf70', fg='#6b3d00', fill='FFCF70',
               p=0.0, pt=136.7, lo=91.7, hit=35,
               troi=111.9, vroi=136.7, vn=61, vhit=35, avg=477),
    'W1': dict(kind='umatan', tier='main', axis='win_prob>=40', order=1,
               label='W1 馬単 軸wp40+×単50騎50×11頭以上',
               short='軸wp40+ × 単50騎50 × 11頭以上',
               bg='#d9e8fb', fg='#1f4e79', fill='D9E8FB',
               p=0.0, pt=122.6, lo=75.9, hit=53,
               troi=122.6, vroi=122.6, vn=202, vhit=53, avg=960),
    'U1': dict(kind='umatan', tier='main', axis='win_prob>=40', order=2,
               label='U1 馬単 軸wp40+×単50騎50',
               short='軸wp40+ × 単50騎50',
               bg='#d9e8fb', fg='#1f4e79', fill='D9E8FB',
               p=0.0743, pt=153.3, lo=63.6, hit=32,
               troi=0.0, vroi=153.3, vn=431, vhit=32, avg=2065),
    # ══════════════════════════════════════════════════════════
    # ★v135_021 (2026/08): axis_enhanced 探索(FDR再判定つき)由来の本線条件を追加。
    #   出自: prob_model_conditions_with_axis.json を Benjamini-Hochberg で再判定。
    #   FDR/Bonferroni では族6433に対し有意0件だが、一貫性(学習&検証>=100%)・
    #   大標本・十分な的中数・クリーンな構造(大穴帯/EV軸を含まない)で選抜した
    #   「グレー採用」条件。実弾は少額に留めること(CI下限は100%未満)。
    #
    #   ★重複回避のため不採用にした候補:
    #     - 軸wp40×単50騎50  = U1 と完全同一。
    #     - 軸wp40×単70騎50  ⊆ U1(単70⊆単50)。先着dedupで新規買い目ゼロ。
    #   ★P1/P2 は S2/F4(複勝確率上位2頭・順位ベース)とは選択ロジックが異なる
    #     閾値ベースで、包含関係を持たない(独自の買い目を追加する)。
    #   ★order は F4(-1)/S2(0)/U1(2) の後の 3/4。既存本線のタグを優先させ、
    #     P1/P2 は紐抜けを埋める側に回す(本線 order の最大は 4 を維持)。
    'P1': dict(kind='umatan', tier='main', axis='win_prob>=50', order=3,
               label='P1 馬単 軸wp50+×複40単60',
               short='軸wp50+ × 複勝率40 & 単指数60',
               bg='#f7d6de', fg='#8a1c3b', fill='F7D6DE',
               p=0.255, pt=140.5, lo=70.6, hit=51,
               troi=116.1, vroi=140.5, vn=200, vhit=51, avg=551),
    'P2': dict(kind='umatan', tier='main', axis='win_prob>=50', order=4,
               label='P2 馬単 軸wp50+×複40人7',
               short='軸wp50+ × 複勝率40 & 単勝人気7',
               bg='#d0eef7', fg='#125a6b', fill='D0EEF7',
               p=0.25, pt=133.6, lo=67.3, hit=59,
               troi=116.8, vroi=133.6, vn=236, vhit=59, avg=535),
    # ══ 紐カバー枠 (tier='cover') ═════════════════════════════════
    #   ★v135_015: 旧「紐カバー枠」を廃止し『紐カバー枠』へ全面変更。
    #
    #   設計原理(被覆率集計で判明した最重要事項):
    #     カバー枠の軸atomは本線枠と同一(win_prob>=40/50)でなければならない。
    #     旧A20/A25(軸wp>=20/25)は軸が緩いため、本線枠が買っていない
    #     レースで単独発火し、発火Rが663→1,421と倍増していた。
    #     これは紐抜けを埋めているのではなく、別物の穴狙いである。
    #     軸を本線と揃えて『相手だけ』を広げることで、初めて紐抜けだけを
    #     埋める構造になる。★軸atomを緩める変更は絶対にしないこと。
    #
    #   ・全条件 馬単。本線枠とは別財布(GZ_ANA_CAP_RATIO)で少点数のみ購入。
    #   ・order は本線(max=4)より後の 13以降。先着優先で本線のタグが残る。
    #
    #   ★E1-E3 は単勝EV帯を相手条件に使う。EV帯の効き方は軸の強さで反転する:
    #       軸wp>=50(堅い) → 低EV帯(0-1)が有効。相手も人気側で決まる。
    #       軸wp>=40       → 高EV帯(1.5+)が有効。相手が荒れる余地がある。
    #     EVを『軸』に使うと壊滅する(プール検証ROI 48.5%)。軸を締めた上での
    #     相手選別にのみ機能する。EV帯の上限を外すと崩れるので緩めないこと。
    'E1': dict(kind='umatan', tier='cover', axis='win_prob>=50', order=13,
               label='E1 馬単 軸wp50+×単60EV0-1',
               short='軸wp50+ × 単指数60 & EV0-1',
               bg='#e6e0f8', fg='#3f3a7a', fill='E6E0F8',
               p=0.1636, pt=140.7, lo=34.6, hit=9,
               troi=0.0, vroi=140.7, vn=55, vhit=9, avg=860),
    'C1': dict(kind='umatan', tier='cover', axis='win_prob>=50', order=14,
               label='C1 馬単 軸wp50+×複20人5',
               short='軸wp50+ × 複勝率20 & 人気5',
               bg='#ead9f5', fg='#5b2c8d', fill='EAD9F5',
               p=0.0945, pt=65.0, lo=26.6, hit=12,
               troi=0.0, vroi=65.0, vn=127, vhit=12, avg=688),
    'C2': dict(kind='umatan', tier='cover', axis='win_prob>=50', order=15,
               label='C2 馬単 軸wp50+×騎40人5',
               short='軸wp50+ × 騎手40 & 人気5',
               bg='#f0e2d0', fg='#7a4a12', fill='F0E2D0',
               p=0.0903, pt=97.8, lo=48.2, hit=13,
               troi=0.0, vroi=97.8, vn=144, vhit=13, avg=1083),
    # ★C3削除(v135_019): 実質29点・取り分率3.9%。U1と91.9%重複。
    'E3': dict(kind='umatan', tier='cover', axis='win_prob>=40', order=17,
               label='E3 馬単 軸wp40+×複30EV1.5+',
               short='軸wp40+ × 複勝率30 & EV1.5+',
               bg='#d7f0ef', fg='#15605d', fill='D7F0EF',
               p=0.0784, pt=99.6, lo=34.1, hit=16,
               troi=0.0, vroi=99.6, vn=204, vhit=16, avg=1270),
    # ★E2削除(v135_019): 114点で的中2本、ROI 18.9%、CI上限53.2%。
    #   累積ROIを唯一明確に下げていた(115.1%→108.6%)。
    #
    # ══════════════════════════════════════════════════════════
    # ★v135_021 (2026/08): 馬連 R3 を紐カバー枠に少額で追加。
    #   出自: axis_enhanced 探索の採用条件⑤(検証ROI127.5% N342 的中74)。
    #   ★これは唯一 kind='umaren' かつ 軸=騎手指数50 の条件。旧R3(kishu軸馬連)の
    #     復活にあたる(_gz_axis_ks50 は既存)。本線枠の『軸は必ずwin_prob』原則の
    #     例外であり、分散目的で cover 枠(別財布・少額)にのみ置く。本線には入れない。
    #   ★相手は 単勝確率_生>=40 & 単勝オッズ 1-10(堅い人気サイド)。馬連なので
    #     順序なし。umaren 収集(_gz_collect ..., 0, False)で frozenset 重複排除。
    'R3': dict(kind='umaren', tier='cover', axis='kishu_idx>=50', order=20,
               label='R3 馬連 軸騎50×単勝率40オッズ1-10',
               short='軸騎手50 × 単勝確率40 & オッズ1-10(馬連)',
               bg='#e4ecc8', fg='#4a5a12', fill='E4ECC8',
               p=0.2164, pt=127.5, lo=67.5, hit=74,
               troi=100.0, vroi=127.5, vn=342, vhit=74, avg=589),
    # ══════════════════════════════════════════════════════════
    # ★v135_028: 本線枠を『軸×紐 勝ち筋マップ』の最良骨格に全面差し替え。
    #   有望条件.json の sign(符号一致)通過14条件のうち、的中率が高く
    #   サンプル数(再現性)も厚い win_prob>=50 × place_prob>=40 系を本線に採用。
    #   ★同じ買い目の重複計上を避けるため、包含関係で最も広い1条件へ集約:
    #     - 馬単 MB1 = 軸wp50 × 複勝率40 は J01/J03/J04/J05/J06/J07/J11 を内包
    #       (それらは全て place40 に単指数/騎手/人気の追加フィルタを掛けた部分集合)。
    #       採用根拠: J06(place40単独) 検ROI127.6 / 的中37.0% / N227 / CI下限91.7、
    #       内包する J01 は 検ROI139.9 / 的中35.9% / N117 / CI下限96.8(最良)。
    #     - 馬連 MB2 = 軸wp50 × 複勝率40 & 騎手50 (=J16 検ROI111.5 / 的中41.9%)。
    #   ※いずれも sign通過だが CI下限は100%未満 → 実弾は損失許容の範囲に留めること。
    # ══════════════════════════════════════════════════════════
    # ══════════════════════════════════════════════════════════
    # ★v135_029: 本線枠を『骨格(軸wp50 × 騎手/勝率)』提案①②③に全面差し替え。
    #   出自: prob_model_conditions_with_axis.json の robust条件(学習ROI>=100 かつ
    #   検証ROI>=100)から、汎用共通骨格の代表3条件を選抜。
    #   ★騎手指数・単勝人気は「補正前(生)」の列で判定する:
    #       騎手指数  → group['騎手指数'](騎手オッズ補正を掛けない生の値)
    #       単勝人気  → group['単勝人気'](adjusted_popularity ではない生の市場人気)
    #     → _partners 実装で _ge(_i,'騎手指数',..) / _pop_le(_i,..) を使うことで担保。
    #   ※いずれも sign等の統計判定は未通過。実弾は損失許容の範囲に留めること。
    #
    #   PA(提案①・本命型/回収重視) 軸wp50 × 騎手40 & 勝率15  検ROI123.2 的中27.8% N250
    #   PB(提案②・的中率重視)      軸wp50 × 勝率15 & 人気2以内 検ROI121.2 的中35.1% N185
    #   PC(提案③・安定/大標本)     軸騎手60 × 複勝率20 & 単指数60 検ROI112.6 的中8.6% N8846
    # ══════════════════════════════════════════════════════════
    'PA': dict(kind='umatan', tier='main', axis='win_prob>=50', order=0,
               label='PA 馬単 軸wp50+×騎手40&勝率15(提案①本命型)',
               short='軸wp50+ × 騎手指数40 & 単勝確率15',
               bg='#9fd8a8', fg='#14532d', fill='9FD8A8',
               p=0.278, pt=123.2, lo=0.0, hit=70,
               troi=120.3, vroi=123.2, vn=250, vhit=70, avg=443),
    'PB': dict(kind='umatan', tier='main', axis='win_prob>=50', order=1,
               label='PB 馬単 軸wp50+×勝率15&人気2以内(提案②的中重視)',
               short='軸wp50+ × 単勝確率15 & 単勝人気2以内',
               bg='#ffcf70', fg='#6b3d00', fill='FFCF70',
               p=0.351, pt=121.2, lo=0.0, hit=65,
               troi=114.6, vroi=121.2, vn=185, vhit=65, avg=345),
    'PC': dict(kind='umatan', tier='main', axis='kishu_idx>=60', order=2,
               label='PC 馬単 軸騎手60+×複勝率20&単指数60(提案③安定型)',
               short='軸騎手60+ × 複勝確率20 & 単指数60',
               bg='#d9e8fb', fg='#1f4e79', fill='D9E8FB',
               p=0.086, pt=112.6, lo=0.0, hit=761,
               troi=100.7, vroi=112.6, vn=8846, vhit=761, avg=1309),

    # ══════════════════════════════════════════════════════════
    # ★v136_001: 最適紐008(278,100通り)の再分析で残った本線枠4条件。
    #
    #   選別のふるい(3段):
    #     (a) サンプル数  : 対象80R以上・的中20点以上
    #     (b) 単調性      : しきい値を1段上げるごとに的中率/回収率が階段状に伸びるか
    #     (c) 近傍安定性  : ソート順4種×上位N5段階の全20通りで回収率の最小値が95%超
    #
    #   (b)の実測(軸=補正後勝率40-49.9%・ワイド):
    #     単勝確率のみ: 2.1→43.7% / 4.6→43.9% / 8.4→45.8% / 15.6→54.0%
    #     単指数のみ  :  41→53.0% /  50→52.2% /  58→52.1% /   68→54.0%
    #     → 両方を課した「勝率15.6 & 単指数68」で 的中54.0% / 回収104.1%(174R)。
    #   (b)の実測(軸=補正後勝率50%+・ワイド):
    #     複勝確率: 8.9→50.2% / 18.3→50.2% / 30.4→51.8% / 47.6→65.5%
    #     → 47.6で明確に折れる。単勝確率8.4や紐馬指数39.8を足しても n=112でほぼ同値
    #       (的中66.1%/回収100.5%)のため、複勝確率47.6が本体で追加項目は無意味。
    #
    #   ★券種は軸の強さで切り替える(同じ買い目でも回収率が40pt動く):
    #     軸40-49.9%: ワイド104.1 / 馬連111.3 / 馬単 82.7  → 馬単は買わない
    #     軸50%+    : ワイド 98.8 / 馬連118.5 / 馬単139.7  → ワイドは買わない
    #                 (ワイドは的中64.9%と最高だが回収98.8%のトリガミ。次点枠で観察)
    #
    #   ★軸・相手とも『補正後』の列で判定する(最適紐008の探索が予想表の表示列で
    #     行われているため)。他条件が使う '_生' 列とは別物なので混同しないこと。
    #   ★騎手指数は一切使わない(日付フィルタ疑い。ヘッダ注記(3)を参照)。
    #   ※いずれもブロックブートストラップCI下限は未算出。実弾は損失許容の範囲で。
    # ══════════════════════════════════════════════════════════
    'HA1': dict(kind='wide', tier='main', axis='adj_win40_49', order=0,
                label='HA1 ワイド 軸wa40-49×勝率15.6&単指数68',
                short='軸[補正後勝率40-49.9%] × 単勝確率15.6 & 単指数68',
                bg='#9fd8a8', fg='#14532d', fill='9FD8A8',
                p=0.540, pt=104.1, lo=0.0, hit=94,
                troi=104.1, vroi=104.1, vn=174, vhit=94, avg=193),
    'HA2': dict(kind='umaren', tier='main', axis='adj_win40_49', order=1,
                label='HA2 馬連 軸wa40-49×勝率15.6&単指数68',
                short='軸[補正後勝率40-49.9%] × 単勝確率15.6 & 単指数68',
                bg='#bfe3c6', fg='#14532d', fill='BFE3C6',
                p=0.299, pt=111.3, lo=0.0, hit=52,
                troi=111.3, vroi=111.3, vn=174, vhit=52, avg=372),
    'HB1': dict(kind='umaren', tier='main', axis='adj_win>=50', order=2,
                label='HB1 馬連 軸wa50+×複勝確率47.6',
                short='軸[補正後勝率50%+] × 複勝確率47.6',
                bg='#ffcf70', fg='#6b3d00', fill='FFCF70',
                p=0.412, pt=118.5, lo=0.0, hit=47,
                troi=118.5, vroi=118.5, vn=114, vhit=47, avg=288),
    'HB2': dict(kind='umatan', tier='main', axis='adj_win>=50', order=3,
                label='HB2 馬単 軸wa50+×複勝確率47.6',
                short='軸[補正後勝率50%+] × 複勝確率47.6',
                bg='#f2a65a', fg='#6b3d00', fill='F2A65A',
                p=0.325, pt=139.7, lo=0.0, hit=37,
                troi=139.7, vroi=139.7, vn=114, vhit=37, avg=430),

    # ══════════════════════════════════════════════════════════
    # ★v136_005: 本線枠を『紐条件アーカイブ』(最適紐016〜026の再分析。馬単/馬連/
    #   ワイド別に的中率と再現性で選抜した9条件)へ全面差し替え。次点本線枠(JITEN)は
    #   廃止(JITEN_ENABLE=False)。ファイル冒頭の changelog(★v136_005)を参照。
    #
    #   ◎=複数ファイルで回収率100%超を一貫して確認 / ○=単一系統だがR>=100で
    #   要因分解の結論と整合 / △=単一ファイル・小標本(R<100)のため要検証。
    #   いずれもこのスクリプト側でCI下限を再算定したものではない。
    # ══════════════════════════════════════════════════════════
    'AR1': dict(kind='umatan', tier='main', axis='adj_win>=50', order=0,
                label='AR1 馬単 軸wa50+×複勝確率47.7(紐条件アーカイブ◎)',
                short='軸[補正後勝率50%+] × 複勝確率47.7',
                bg='#9fd8a8', fg='#14532d', fill='9FD8A8',
                p=0.0, pt=128.9, lo=0.0, hit=30,
                troi=0.0, vroi=128.9, vn=134, vhit=40, avg=0),
    'AR2': dict(kind='umatan', tier='main', axis='tan_idx>=90', order=1,
                label='AR2 馬単 軸単指数90+×複勝確率47.7(紐条件アーカイブ○)',
                short='軸[単指数90+] × 複勝確率47.7',
                bg='#ffcf70', fg='#6b3d00', fill='FFCF70',
                p=0.0, pt=119.2, lo=0.0, hit=13,
                troi=0.0, vroi=119.2, vn=216, vhit=27, avg=0),
    'AR3': dict(kind='umatan', tier='main', axis='adj_win>=50', order=2,
                label='AR3 馬単 軸wa50+×複勝確率47.7&騎手補正48.1(紐条件アーカイブ△要検証)',
                short='軸[補正後勝率50%+] × 複勝確率47.7 & 騎手補正48.1',
                bg='#d9e8fb', fg='#1f4e79', fill='D9E8FB',
                p=0.0, pt=159.6, lo=0.0, hit=39,
                troi=0.0, vroi=159.6, vn=77, vhit=30, avg=0),
    'AR4': dict(kind='umaren', tier='main', axis='adj_win>=50', order=0,
                label='AR4 馬連 軸wa50+×複勝確率47.7(紐条件アーカイブ◎)',
                short='軸[補正後勝率50%+] × 複勝確率47.7',
                bg='#9fd8a8', fg='#14532d', fill='9FD8A8',
                p=0.0, pt=114.3, lo=0.0, hit=39,
                troi=0.0, vroi=114.3, vn=134, vhit=52, avg=0),
    'AR5': dict(kind='umaren', tier='main', axis='tan_idx>=90&pop<=1', order=5,
                label='AR5 馬連 軸単指数90+&人気1以内×複勝確率47.7(第2版○維持)',
                short='軸[単指数90+ & 人気1以内] × 複勝確率47.7',
                bg='#ffcf70', fg='#6b3d00', fill='FFCF70',
                p=0.0, pt=115.0, lo=0.0, hit=31,
                troi=0.0, vroi=115.0, vn=113, vhit=35, avg=0),
    'AR6': dict(kind='umaren', tier='main', axis='tan_idx>=90&pop<=1', order=6,
                label='AR6 馬連 軸単指数90+&人気1以内×勝率15.6&単指数68(第2版△維持・R56)',
                short='軸[単指数90+ & 人気1以内] × 単勝確率15.6 & 単指数68',
                bg='#d9e8fb', fg='#1f4e79', fill='D9E8FB',
                p=0.0, pt=122.9, lo=0.0, hit=38,
                troi=0.0, vroi=122.9, vn=56, vhit=21, avg=0),
    'AR7': dict(kind='wide', tier='main', axis='tan_idx>=90&pop<=1', order=2,
                label='AR7 ワイド 軸単指数90+&人気1以内×勝率15.6&単指数68(第2版△維持・R56)',
                short='軸[単指数90+ & 人気1以内] × 単勝確率15.6 & 単指数68',
                bg='#9fd8a8', fg='#14532d', fill='9FD8A8',
                p=0.0, pt=125.4, lo=0.0, hit=64,
                troi=0.0, vroi=125.4, vn=56, vhit=36, avg=0),
    'AR8': dict(kind='wide', tier='main', axis='tan_idx>=90&pop<=1&odds<=3', order=1,
                label='AR8 ワイド 軸単指数90+&人気1以内&オッズ3以内×複勝確率47.7(紐条件アーカイブ○)',
                short='軸[単指数90+ & 人気1以内 & オッズ3以内] × 複勝確率47.7',
                bg='#ffcf70', fg='#6b3d00', fill='FFCF70',
                p=0.0, pt=101.5, lo=0.0, hit=54,
                troi=0.0, vroi=101.5, vn=107, vhit=58, avg=0),
    'AR9': dict(kind='wide', tier='main', axis='adj_win>=50', order=2,
                label='AR9 ワイド 軸wa50+×単指数68&騎手補正58.9(紐条件アーカイブ△要検証)',
                short='軸[補正後勝率50%+] × 単指数68 & 騎手補正58.9',
                bg='#d9e8fb', fg='#1f4e79', fill='D9E8FB',
                p=0.0, pt=122.5, lo=0.0, hit=57,
                troi=0.0, vroi=122.5, vn=40, vhit=23, avg=0),
    # ══════════════════════════════════════════════════════════
    # ★v136_006: 本線枠を『紐条件アーカイブ第2版』(最適紐001・016〜026を統合し、
    #   閾値近傍・軸近傍・余裕的中数・001以後の追加レースで再検証)に合わせて組み替え。
    #   ◎=R100以上・閾値/軸の近傍とも安定 / ○=R70以上で同 / △=近傍の一部が弱い・R<70。
    #   紐はすべて『条件該当馬のうち紐馬指数最大の1頭』(N=1)。
    #   ★数値は探索期間内の実績。001以後に追加された25〜35Rでは全条件の水準が
    #     概ね半分に落ちており(小標本)、将来の回収を保証しない。
    # ══════════════════════════════════════════════════════════
    'V1': dict(kind='umaren', tier='main', axis='tan_idx>=90&pop<=1', order=0,
                label='V1 馬連 軸単指数90+&人気1以内×単勝確率8.4(第2版◎)',
                short='軸[単指数90+ & 人気1以内] × 単勝確率8.4',
                bg='#9fd8a8', fg='#14532d', fill='9FD8A8',
                p=0.0, pt=122.4, lo=0.0, hit=26,
                troi=0.0, vroi=122.4, vn=174, vhit=45, avg=0),
    'V2': dict(kind='umaren', tier='main', axis='tan_idx>=90&pop<=1&odds<=3', order=1,
                label='V2 馬連 軸単指数90+&人気1以内&オッズ3以内×複勝確率47.7&単指数50(第2版◎)',
                short='軸[単指数90+ & 人気1以内 & オッズ3以内] × 複勝確率47.7 & 単指数50',
                bg='#9fd8a8', fg='#14532d', fill='9FD8A8',
                p=0.0, pt=124.6, lo=0.0, hit=33,
                troi=0.0, vroi=124.6, vn=103, vhit=34, avg=0),
    'V3': dict(kind='umaren', tier='main', axis='tan_idx>=90&pop<=1&odds<=3', order=2,
                label='V3 馬連 軸単指数90+&人気1以内&オッズ3以内×単勝確率8.4&単指数68(第2版◎)',
                short='軸[単指数90+ & 人気1以内 & オッズ3以内] × 単勝確率8.4 & 単指数68',
                bg='#9fd8a8', fg='#14532d', fill='9FD8A8',
                p=0.0, pt=126.1, lo=0.0, hit=28,
                troi=0.0, vroi=126.1, vn=120, vhit=34, avg=0),
    'V4': dict(kind='umaren', tier='main', axis='tan_idx>=90&pop<=1&odds<=2', order=3,
                label='V4 馬連 軸単指数90+&人気1以内&オッズ2以内×単勝確率4.6(第2版◎)',
                short='軸[単指数90+ & 人気1以内 & オッズ2以内] × 単勝確率4.6',
                bg='#9fd8a8', fg='#14532d', fill='9FD8A8',
                p=0.0, pt=133.4, lo=0.0, hit=30,
                troi=0.0, vroi=133.4, vn=104, vhit=31, avg=0),
    'V5': dict(kind='umaren', tier='main', axis='tan_idx>=90&pop<=1&odds<=3', order=4,
                label='V5 馬連 軸単指数90+&人気1以内&オッズ3以内×複勝確率47.7&単指数68(第2版○)',
                short='軸[単指数90+ & 人気1以内 & オッズ3以内] × 複勝確率47.7 & 単指数68',
                bg='#ffcf70', fg='#6b3d00', fill='FFCF70',
                p=0.0, pt=140.4, lo=0.0, hit=35,
                troi=0.0, vroi=140.4, vn=71, vhit=25, avg=0),
    'V6': dict(kind='wide', tier='main', axis='tan_idx>=90&pop<=1&odds<=2', order=0,
                label='V6 ワイド 軸単指数90+&人気1以内&オッズ2以内×複勝確率47.7(第2版○・AR8置換)',
                short='軸[単指数90+ & 人気1以内 & オッズ2以内] × 複勝確率47.7',
                bg='#ffcf70', fg='#6b3d00', fill='FFCF70',
                p=0.0, pt=113.0, lo=0.0, hit=65,
                troi=0.0, vroi=113.0, vn=71, vhit=46, avg=0),
    'V7': dict(kind='wide', tier='main', axis='adj_win>=50&tan_idx>=90&pop<=1&odds<=3', order=1,
                label='V7 ワイド 軸wa50+&単指数90+&人気1以内&オッズ3以内×単勝確率4.6&複指数20(第2版○)',
                short='軸[補正後勝率50%+ & 単指数90+ & 人気1以内 & オッズ3以内] × 単勝確率4.6 & 複指数20',
                bg='#ffcf70', fg='#6b3d00', fill='FFCF70',
                p=0.0, pt=120.8, lo=0.0, hit=55,
                troi=0.0, vroi=120.8, vn=73, vhit=40, avg=0),
    # ══════════════════════════════════════════════════════════
    # ★v136_018: 本線枠 = BADO 実戦条件シートの実戦候補3条件(紐条件探索3・4)。
    #   数値は 2026/7/16〜9/23 の予想表(9/24再生成)×月次配当CSV での成績。
    #   全期間を見て選んだ条件のため上振れを含む(選び方の検証では実力は回収100〜110%程度)。
    # ══════════════════════════════════════════════════════════
    'UR1': dict(kind='umaren', tier='main', axis='S:ti90max&odds<1.95', order=0,
                label='UR1 馬連① 軸[単指数90+の最大&予想オッズ1.9以下]×単勝確率8以上の紐馬指数1位(実戦シート)',
                short='馬連① 軸[単指数90+最大 & 予想オッズ1.9以下] × 単勝確率8以上→紐馬指数1位',
                bg='#9fd8a8', fg='#14532d', fill='9FD8A8',
                p=0.0, pt=121.3, lo=0.0, hit=30,
                troi=121.6, vroi=120.9, vn=99, vhit=30, avg=400),
    'UR2': dict(kind='umaren', tier='main', axis='S:ti90-93max&odds<1.45', order=1,
                label='UR2 馬連② 軸[単指数90〜93&予想オッズ1.4以下]×人気3以内&紐馬指数50以上の紐馬指数1位(実戦シート)',
                short='馬連② 軸[単指数90〜93 & 予想オッズ1.4以下] × 人気3以内&紐馬指数50以上→紐馬指数1位',
                bg='#7cc98b', fg='#0f3d21', fill='7CC98B',
                p=0.0, pt=156.2, lo=0.0, hit=49,
                troi=165.5, vroi=146.3, vn=39, vhit=19, avg=320),
    'WD1': dict(kind='wide', tier='main', axis='S:aw50max&ti>=90&pop1', order=0,
                label='WD1 ワイド① 軸[単勝確率50%+の最大&単指数90+&人気1]×複勝確率45以上&単指数65以上の紐馬指数1位(実戦シート)',
                short='ワイド① 軸[単勝確率50%+最大 & 単指数90+ & 人気1] × 複勝確率45以上&単指数65以上→紐馬指数1位',
                bg='#ffcf70', fg='#6b3d00', fill='FFCF70',
                p=0.0, pt=132.4, lo=0.0, hit=85,
                troi=135.5, vroi=127.7, vn=33, vhit=28, avg=156),
}

# ══════════════════════════════════════════════════════════
# ★v135_022 (2026/08): 本線枠を U1 / S2 の2条件に縮小。
#
#   根拠(axis_enhanced 探索 + ランク族探索の再分析):
#     ・P1(複40単60)/P2(複40人7) は place_prob>=40 単独と実質同一集合。
#       P2 の pop<=7 は1点も削っていない(159R/1.50点/R → 157R/1.50点/R)。
#       P1 の tan_idx>=60 も学習マージンを +4.0→+3.9 と微減させるだけ。
#       3条件を並べると同じ買い目の三重計上になる。
#     ・F4 の EV上位フィルタは的中率を破壊する(S2 41.5% → F4 12.5%)。
#       軸を横断しても検証ROIの崩壊が 8/8。反転版(EV下位)も検74.9%で不可。
#     ・E1/C1/C2/E3/R3(旧紐カバー枠)は JSON駆動の参考枠へ移管。
#
#   ★削除ではなくフィルタで縮小している。定義本体と _partners_* 実装を
#     残すことで、条件を戻す判断をしたとき GZ_MAIN_KEEP の1行で復帰できる。
#     未使用の実装が残るが、レジストリを直接削るより事故が少ない。
# ══════════════════════════════════════════════════════════
# ★v135_028: 本線枠を『軸×紐 勝ち筋マップ』集約2条件 MB1(馬単)/MB2(馬連) に差し替え。
#   旧 A1/W1/U1/P1/P2 等の定義は残置(この集合に名前を戻せば復帰可能)。
#   ※次点本線枠(旧 subline A/B/C)は別途 SUBLINE_LEGACY_ENABLE=False で停止し、
#     sign通過の残条件を『次点本線枠(勝ち筋マップ)』として表示専用で再構成する。
# ★v136_001: 本線枠を最適紐008の提案4条件へ全面差し替え(PA/PB/PC は残置=名前を戻せば復帰)。
# ★v136_005: 本線枠を『紐条件アーカイブ』(最適紐016〜026)の提案9条件へ全面差し替え。
#   HA1/HA2/HB1/HB2 は定義を残置(この集合に名前を戻せば復帰可能)。
# ★v136_006: 紐条件アーカイブ第2版の再判定に合わせて組み替え。
#   維持 AR5(○)/AR6(△)/AR7(△)、追加 V1〜V7。
#   外した AR1/AR3/AR4(観察)・AR2/AR8/AR9(除外)は定義を残置(ここへ名前を戻せば復帰)。
#   馬単は本線枠なし(騎手指数を使わない条件で基準を通るものが無かったため)。
# ★v136_018: 旧本線枠(V1〜V7・AR5〜AR7)を廃止し、実戦条件シートの3条件だけにする。
GZ_MAIN_KEEP = {'UR1', 'UR2',   # 馬連
                'WD1'}          # ワイド
# ★v135_022: 補充馬連(買い目0点レースの穴埋め馬連3点)の有効/無効。
#   True に戻すと従来どおり『紐カバー枠』として馬連が表示される。
REF_UMAREN_ENABLE = False
_GZ_COND_DEFS_FULL = dict(GZ_COND_DEFS)
GZ_COND_DEFS = {k: v for k, v in GZ_COND_DEFS.items()
                if k in GZ_MAIN_KEEP}
for _k in GZ_COND_DEFS:
    GZ_COND_DEFS[_k]['tier'] = 'main'
assert set(GZ_COND_DEFS) == GZ_MAIN_KEEP, (
    f'本線枠の条件が欠けています: {GZ_MAIN_KEEP - set(GZ_COND_DEFS)}')

# ★v135_011: 枠(tier)の表示名と、枠判定ヘルパ。
GZ_TIER_LABEL = {'main': '本線枠', 'cover': '紐カバー枠'}
# ★v136_001: 券種の表示名(ワイドの本線枠採用に伴い辞書化)。
GZ_KIND_JP = {'umatan': '馬単', 'umaren': '馬連', 'wide': 'ワイド'}
GZ_KIND_SEP = {'umatan': '→', 'umaren': '-', 'wide': '-'}

# ══════════════════════════════════════════════════════════
# ★v136_004: 参考予想 — 軸=◎(アンサンブル単勝確率1位) の買い目
#
#   ・旧・参考枠(REF_WP1 / gz_reference.py + 有望条件.json 駆動)は全廃した。
#   ・相手 = 『複勝確率(補正後) HIMO_PLACE_PROB_MIN%以上 または
#     紐馬指数 HIMO_HIMOBA_IDX_MIN以上』の該当馬“全員”(_himo_all_idxs)。
#     該当0頭のときのみ紐馬指数上位3頭にフォールバック。
#     ★予想印▲△☆(_himo_partner_idxs)は記号が3種類しかないため上位3頭までの
#       表示だが、参考予想の相手選定はそれとは独立して該当馬全員を対象にする。
#   ・券種は馬連・ワイドの2種(REF_HIMO_BET_TYPESで変更可)。
#   ・参考予想は表示専用。実弾ではなく gz_stakes / 財布ロジックを一切通らない。
# ══════════════════════════════════════════════════════════
REF_HIMO_BET_TYPES = ('umaren', 'wide')

# ★v135_025: 次点本線枠A — 相手に要求する 複勝確率_生 の下限(%)
#   40-45%帯が実測で最も悪く(馬連ROI 37%)、50%以上で明確に分かれたため50に設定。
SUBLINE_PARTNER_MIN_PLACE = 50.0

# ★v135_026: 次点本線枠B — 軸wp40+ × 相手 複指数>=24 × 10頭以下
#   8/1-8/23 実測(軸wp40+ 238R のうち10頭以下167R、複指数24+の相手は60Rで74点):
#     馬単 74点 的中13.5% ROI181.4% / 馬連 14.9% ROI161.8% / ワイド 31.1% ROI111.6%
#   同条件を11頭以上(=W1が採用する側)で見ると馬単37点 ROI73.0% と逆転するため、
#   頭数上限を必須とする。※N=74、CI下限は100%未満(馬単47%/ワイド65%)。
SUBLINE_B_AXIS_MIN_WP = 40.0      # 軸の 単勝確率_生 下限(%)
SUBLINE_B_PARTNER_MIN_FUKU = 24.0  # 相手の 複指数 下限
SUBLINE_B_MAX_FIELD = 10           # 出走頭数の上限(これ以下で発火)

# ★v135_027: 次点本線枠C — 有望条件.json の最良骨格(J01/J16系)を実装。
#   軸 = 単勝確率_生>=50 の最上位1頭、相手 = 軸以外で 複勝確率_生>=40 かつ
#   騎手指数>=50 の全馬。馬単(軸→相手)と馬連で流す。
#   参考: 有望条件.json J01(馬単 軸wp50×place40&kishu50) valid_ROI139.9/的中35.9%、
#         J16(馬連 同骨格) valid_ROI111.5/的中41.9%。sign通過だが CI下限は100%未満のため
#         『検証中の候補』位置づけ(本線枠には昇格させない)。
SUBLINE_C_AXIS_MIN_WP = 50.0       # 軸の 単勝確率_生 下限(%)
SUBLINE_C_PARTNER_MIN_PLACE = 40.0  # 相手の 複勝確率_生 下限(%)
SUBLINE_C_PARTNER_MIN_KISHU = 50.0  # 相手の 騎手指数 下限

# ★v135_028: 旧・次点本線枠(A/B/C)を全停止。以降は『次点本線枠(勝ち筋マップ)』へ一本化。
#   本線枠(実弾)= MB1/MB2 に採用しなかった sign通過条件を、表示専用の次点枠として再構成する。
#   Trueに戻せば旧A/B/Cが復活する(定義・生成コードは残置)。
SUBLINE_LEGACY_ENABLE = False

# ★v136_005: 次点本線枠(勝ち筋マップ/JITEN_MAP_DEFS)を廃止。
#   本線枠を『紐条件アーカイブ』9条件へ全面差し替えたことに伴い、予想表(HTML/Excel/CSV)
#   の【次点本線枠】表示を止める。JITEN_MAP_DEFS・ローダ・生成コードは残置し、
#   True に戻せば従来どおり復活する(定義を消すよりレジストリ不整合の事故が少ない)。
JITEN_ENABLE = True   # ★v136_018: 実戦条件シートの『観察のみ』2条件で再開

# ★v135_028: 次点本線枠(勝ち筋マップ)— sign通過14条件のうち本線(MB1/MB2)に
#   採用しなかった残条件。表示専用(実弾ではない/賭け金・財布を通らない)。
#   ★同じ買い目は本線・次点を通して1つに集約(本線を優先し次点から重複を除く)。
#   ★CI下限は全て100%未満のため『検証中の候補』。ROI降順で優先度付け。
#   [tag, kind, axisの種別, 相手アトム, 表示ラベル, 検証統計コメント]
#   ・axis: 'w40'=軸wp40+ / 'w50'=軸wp50+ / 'ao80_25'=軸[単指数80+&オッズ2.5以下]
# ★v135_030: 次点本線枠を『有力条件_絞り込み(sign一致 かつ 両期間ROI>=100)』57件から
#   次点枠に適合する条件へ全面差し替え。適合＝軸が w40/w50、相手アトムが
#   複勝確率_生/単指数/騎手指数/単勝確率_生/人気(_pop) のみ(複勝上位N・オッズ・EVは非対応)。
#   ★騎手指数・単勝人気(_pop)は補正前(生)で判定する(_ge/_pop_le)。
#   ★同義の入れ子(軸wp50×勝率15の各サブセット)は避け、相手指標が異なる/標本・的中に
#     幅がある/軸を w50・w40 に分散した6件を選抜。並びは具体条件→広い受け皿(J12)の順。
#   ・axis: 'w40'=軸wp40+ / 'w50'=軸wp50+ / 'ao80_25'=軸[単指数80+&オッズ2.5以下]
# ★v135_032: 次点本線枠は外部JSON(有望条件.json)から読み込む(下記ローダ)。
#   下の _JITEN_MAP_FALLBACK は JSON が無い/壊れている場合のフォールバック定義。
_JITEN_MAP_FALLBACK = [
    # J01: 馬単 軸[単90&オッズ1.8] × 騎手50 & 複勝上位2 (検ROI142.7 / 的中28.6% / N304)
    dict(tag='J01', kind='umatan', axis='ao90_18',
         atoms=[('騎手指数', 50.0, '騎手50'), ('_place_rank', 2.0, '複勝上位2')],
         roi=142.7, hit=28.6, n=304),
    # J02: 馬単 軸[単90&オッズ2.0] × 騎手50 & 複勝上位2 (検ROI130.5 / 的中27.9% / N359)
    dict(tag='J02', kind='umatan', axis='ao90_20',
         atoms=[('騎手指数', 50.0, '騎手50'), ('_place_rank', 2.0, '複勝上位2')],
         roi=130.5, hit=27.9, n=359),
    # J06: 馬単 軸wp50 × 勝率15 & 人気2以内 (検ROI119.9 / 的中35.6% / N194 / CI下限83.6)
    dict(tag='J06', kind='umatan', axis='w50',
         atoms=[('単勝確率_生', 15.0, '勝率15'), ('_pop', 2.0, '人気2')],
         roi=119.9, hit=35.6, n=194),
    # J07: 馬単 軸wp50 × 単指数60 & 勝率15 (検ROI118.7 / 的中29.5% / N292 / CI下限86.0=再現性最良)
    dict(tag='J07', kind='umatan', axis='w50',
         atoms=[('単指数', 60.0, '単指数60'), ('単勝確率_生', 15.0, '勝率15')],
         roi=118.7, hit=29.5, n=292),
    # J18: 馬単 軸[単90&オッズ1.5] × 単指数60 & 複勝上位2 (検ROI108.7 / 的中40.2% / N219)
    dict(tag='J18', kind='umatan', axis='ao90_15',
         atoms=[('単指数', 60.0, '単指数60'), ('_place_rank', 2.0, '複勝上位2')],
         roi=108.7, hit=40.2, n=219),
    # J21: 馬単 軸wp50 × 勝率10 & 人気2以内 (検ROI107.3 / 的中31.5% / N422=大標本 / CI下限86.7)
    dict(tag='J21', kind='umatan', axis='w50',
         atoms=[('単勝確率_生', 10.0, '勝率10'), ('_pop', 2.0, '人気2')],
         roi=107.3, hit=31.5, n=422),
    # J28: 馬単 軸wp50 × 複勝40 & 勝率15 (検ROI105.7 / 的中29.5% / N292)
    dict(tag='J28', kind='umatan', axis='w50',
         atoms=[('複勝確率_生', 40.0, '複勝40'), ('単勝確率_生', 15.0, '勝率15')],
         roi=105.7, hit=29.5, n=292),
    # J33: 馬単 軸wp40 × 単指数70 & 騎手50 (検ROI104.5 / 的中15.8% / N701 / 軸分散)
    dict(tag='J33', kind='umatan', axis='w40',
         atoms=[('単指数', 70.0, '単指数70'), ('騎手指数', 50.0, '騎手50')],
         roi=104.5, hit=15.8, n=701),
    # J57: 馬連 軸[単90&オッズ1.5] × 単指数60 & 複勝上位2 (検ROI101.4 / 的中46.1% / N219 / 券種分散)
    dict(tag='J57', kind='umaren', axis='ao90_15',
         atoms=[('単指数', 60.0, '単指数60'), ('_place_rank', 2.0, '複勝上位2')],
         roi=101.4, hit=46.1, n=219),
    # J12: 馬単 軸wp50 × 勝率15 (検ROI113.5 / 的中28.9% / N311 / 広い受け皿=最後に集約)
    dict(tag='J12', kind='umatan', axis='w50',
         atoms=[('単勝確率_生', 15.0, '勝率15')],
         roi=113.5, hit=28.9, n=311),
]


# ══════════════════════════════════════════════════════════
# ★v135_032: 次点本線枠(表示専用)を外部JSON『有望条件.json』から読み込む。
#   ・既定パスは本スクリプトと同じディレクトリの 有望条件.json。
#     環境変数 JITEN_JSON でパスを上書き可能。
#   ・JSONの各条件を次点枠エンジンが扱える形へ変換する:
#       axis_spec ['ao', t, o]        → 'ao{t}_{o*10}'  (単指数t & オッズo以下)
#       axis_spec ['atom','win_prob',n] → 'w{n}' / place_prob→'pp{n}' / kishu_idx→'ks{n}'
#       partner_atoms [metric, thr]   → (列名, 閾値, ラベル)
#         win_prob→単勝確率_生 / place_prob→複勝確率_生 / tan_idx→単指数 /
#         kishu_idx→騎手指数(補正前) / pop_max→_pop(補正前人気) / place_rank→_place_rank
#   ・次点エンジンが未対応の軸(tan_idx単独等)・相手アトム(odds/ev)を含む条件は自動スキップ。
#   ・優先度は検証ROI降順(先着dedupで高ROIタグが残る)。
#   ・読込失敗/0件のときは _JITEN_MAP_FALLBACK を使用(黙って落とさない)。
# ══════════════════════════════════════════════════════════
JITEN_JSON_PATH = _os.environ.get(
    'JITEN_JSON',
    _os.path.join(_os.path.dirname(_os.path.abspath(__file__)), '有望条件.json'))

# ★v135_034: 次点本線枠は『有望条件のうち検証的中率がこの値以上』のみ採用する。
#   これ未満の条件は次点に載せず、参考枠(gz_reference)側に回す(向こうも同じ閾値で分担)。
#   本線枠(PA)はこの分割の対象外で従来どおり維持。
JITEN_MIN_HIT = float(_os.environ.get('JITEN_MIN_HIT', '25.0'))


def _jiten_axis_code_from_spec(axis_spec, axis_str=None):
    """JSONの axis_spec を次点枠の軸コードへ変換。未対応は None。"""
    try:
        if axis_spec and axis_spec[0] == 'ao':
            _t = int(round(float(axis_spec[1])))
            _o = int(round(float(axis_spec[2]) * 10))
            return f'ao{_t}_{_o}'
        if axis_spec and axis_spec[0] == 'atom':
            _var = axis_spec[1]; _thr = int(round(float(axis_spec[2])))
            _pre = {'win_prob': 'w', 'place_prob': 'pp', 'kishu_idx': 'ks'}.get(_var)
            if _pre:
                return f'{_pre}{_thr}'
    except Exception:
        return None
    return None


def _jiten_atom_from_pair(metric, thr):
    """JSONの相手アトム [metric, thr] を (列名, 閾値, ラベル) へ変換。未対応は None。"""
    try:
        _t = float(thr)
    except Exception:
        return None
    if metric == 'place_rank':
        return ('_place_rank', _t, f'複勝上位{int(_t)}')
    if metric == 'pop_max':
        return ('_pop', _t, f'人気{int(_t)}')
    _col = {'win_prob': '単勝確率_生', 'place_prob': '複勝確率_生',
            'tan_idx': '単指数', 'kishu_idx': '騎手指数'}.get(metric)
    if _col is None:
        return None   # odds_band / ev_band などは次点エンジン未対応
    _lab = {'win_prob': '勝率', 'place_prob': '複勝',
            'tan_idx': '単指数', 'kishu_idx': '騎手'}[metric]
    return (_col, _t, f'{_lab}{int(_t)}')


def _load_jiten_defs_from_json(path):
    with open(path, encoding='utf-8') as _f:
        _d = _json.load(_f)
    _out = []
    for _c in _d.get('conditions', []):
        _kind = _c.get('kind')
        if _kind not in ('umatan', 'umaren', 'wide'):
            continue
        _axc = _jiten_axis_code_from_spec(_c.get('axis_spec'), _c.get('axis'))
        if _axc is None:
            continue
        _atoms = []; _ok = True
        for _pa in (_c.get('partner_atoms') or []):
            _a = _jiten_atom_from_pair(_pa[0], _pa[1])
            if _a is None:
                _ok = False; break
            _atoms.append(_a)
        if not _ok or not _atoms:
            continue
        _st = _c.get('stats', {})
        _hit = float(_st.get('valid_hit_rate', 0.0))
        if _hit < JITEN_MIN_HIT:           # ★的中率25%未満は次点に載せない(参考枠へ)
            continue
        _out.append(dict(
            tag=_c.get('name', '?'), kind=_kind, axis=_axc, atoms=_atoms,
            roi=float(_st.get('valid_roi', 0.0)),
            hit=_hit,
            n=int(_st.get('n_races', 0))))
    _out.sort(key=lambda _x: -_x['roi'])   # 検証ROI降順=優先度
    return _out


try:
    JITEN_MAP_DEFS = _load_jiten_defs_from_json(JITEN_JSON_PATH)
    if not JITEN_MAP_DEFS:
        raise ValueError('次点枠に採用できる条件が0件')
    print(f'[次点本線枠] {JITEN_JSON_PATH} から {len(JITEN_MAP_DEFS)} 件を読み込みました。')
except Exception as _e:
    # フォールバックにも同じ的中率フィルタを適用(次点=的中率25%以上を維持)。
    JITEN_MAP_DEFS = [_d for _d in _JITEN_MAP_FALLBACK
                      if float(_d.get('hit', 0.0)) >= JITEN_MIN_HIT]
    print(f'[次点本線枠] JSON読込に失敗({type(_e).__name__}: {_e})。'
          f'内蔵フォールバック {len(JITEN_MAP_DEFS)} 件を使用します。')


# ══════════════════════════════════════════════════════════
# ★v135_035: 次点本線枠に『おすすめ条件ランキング1～30位』(最適紐007)を追加。
#   出典: 最適紐007_ランキング付き.xlsx シート「おすすめ条件ランキング」。
#   紐馬探索結果(全1,524条件)から 対象50R以上・回収率100%以上 に絞り、
#   券種内の「的中率順位(百分位)」と「回収率順位(百分位)」の相乗平均で
#   スコア化した上位30件(rank=順位)。
#   ・軸/相手の指標は補正後の列で判定する:
#       単勝確率→単勝確率(補正後) / 複勝確率→複勝確率(補正後) / 単指数 / 複指数 / 紐馬指数
#   ・軸『単勝確率50%以上』→'wa50'(補正後 単勝確率>=50)、
#     『単指数90以上』→'ti90'(単指数のみ・オッズ足切りなし)。
#     'wa{n}'(補正後勝率軸)・'ti{n}'(単指数軸)は _jiten_axis に追加実装。
#   ・的中率フィルタ(JITEN_MIN_HIT)は通さず、30件すべてを次点枠へ載せる(指定どおり)。
#   ・stats はランキングの 回収率(roi)/的中率(hit)/対象レース数(n)。
#   ・JITEN_MAP_DEFS の先頭に連結し、重複買い目はおすすめ側(高順位)を優先集約する。
# ══════════════════════════════════════════════════════════
#   ★補正後の列で判定する: 単勝確率→単勝確率(補正後) / 複勝確率→複勝確率(補正後)。
#     ('_生' ではなく) 表示・判定用の補正後列を使う。単指数・複指数・紐馬指数は
#     補正概念のない基礎指数なのでそのまま。
_OSU_ATOM = {
    'win':   ('単勝確率', '勝率'),
    'place': ('複勝確率', '複勝'),
    'tan':   ('単指数', '単指数'),
    'fuku':  ('複指数', '複指数'),
    'himo':  ('紐馬指数', '紐馬'),
}
# [rank, kind, axis, [(metric, thr), ...], roi, hit, n, タイプ]
_OSUSUME_RANK_RAW = [
    (1,  'wide',   'wa50', [('win', 10.2), ('place', 52.1)], 107.9, 76.3, 71, '高的中・堅実'),
    (2,  'wide',   'wa50', [('place', 52.1), ('tan', 61)],   109.2, 75.4, 63, '高的中・堅実'),
    (3,  'wide',   'wa50', [('place', 52.1), ('himo', 37.6)], 106.5, 75.3, 72, '高的中・堅実'),
    (4,  'wide',   'wa50', [('place', 52.1), ('himo', 47.3)], 106.4, 75.0, 72, '高的中・堅実'),
    (5,  'umaren', 'wa50', [('win', 10.2), ('place', 52.1)], 128.3, 51.3, 71, 'バランス'),
    (6,  'umaren', 'wa50', [('place', 52.1), ('himo', 37.6)], 126.6, 50.6, 72, 'バランス'),
    (7,  'umatan', 'wa50', [('place', 52.1), ('tan', 54)],   134.2, 36.8, 72, '高回収・一撃'),
    (8,  'umatan', 'wa50', [('win', 10.2), ('place', 52.1)], 134.2, 36.8, 71, '高回収・一撃'),
    (9,  'umaren', 'wa50', [('place', 52.1), ('tan', 54)],   126.2, 50.0, 72, 'バランス'),
    (10, 'wide',   'wa50', [('place', 52.1)],                105.1, 74.4, 73, '高的中・堅実'),
    (11, 'umaren', 'wa50', [('place', 52.1)],                125.0, 50.0, 73, 'バランス'),
    (12, 'umatan', 'wa50', [('place', 52.1), ('tan', 61)],   139.4, 35.4, 63, '高回収・一撃'),
    (13, 'wide',   'wa50', [('place', 52.1), ('tan', 47)],   105.1, 74.0, 73, '高的中・堅実'),
    (14, 'wide',   'wa50', [('place', 52.1), ('himo', 63.2)], 105.4, 73.1, 50, '高的中・堅実'),
    (15, 'umatan', 'wa50', [('place', 52.1), ('tan', 47)],   132.5, 36.4, 73, '高回収・一撃'),
    (16, 'umatan', 'wa50', [('place', 52.1), ('himo', 37.6)], 132.5, 36.4, 72, '高回収・一撃'),
    (17, 'umaren', 'wa50', [('place', 52.1), ('tan', 61)],   127.5, 47.7, 63, 'バランス'),
    (18, 'wide',   'wa50', [('place', 52.1), ('tan', 54)],   104.5, 73.7, 72, '高的中・堅実'),
    (19, 'wide',   'wa50', [('place', 52.1), ('fuku', 19)],  103.9, 74.3, 66, '高的中・堅実'),
    (20, 'umaren', 'wa50', [('place', 52.1), ('tan', 47)],   124.5, 49.4, 73, 'バランス'),
    (21, 'wide',   'wa50', [('place', 52.1), ('himo', 55.2)], 104.0, 73.5, 65, '高的中・堅実'),
    (22, 'umatan', 'wa50', [('place', 52.1)],                130.8, 35.9, 73, '高回収・一撃'),
    (23, 'umaren', 'wa50', [('place', 52.1), ('himo', 55.2)], 122.1, 48.5, 65, 'バランス'),
    (24, 'wide',   'wa50', [('place', 52.1), ('fuku', 20)],  102.3, 71.7, 57, '高的中・堅実'),
    (25, 'umaren', 'wa50', [('place', 52.1), ('himo', 63.2)], 121.7, 44.2, 50, 'バランス'),
    (26, 'umaren', 'wa50', [('place', 52.1), ('fuku', 19)],  115.9, 48.6, 66, 'バランス'),
    (27, 'umaren', 'wa50', [('place', 52.1), ('fuku', 20)],  114.8, 45.0, 57, 'バランス'),
    (28, 'umatan', 'wa50', [('place', 52.1), ('himo', 55.2)], 125.9, 33.8, 65, '高回収・一撃'),
    (29, 'umaren', 'ti90', [('win', 17.6), ('fuku', 22)],    139.3, 26.8, 54, 'バランス'),
    (30, 'umaren', 'wa50', [('win', 17.6)],                  104.4, 52.7, 55, 'バランス'),
]


def _build_osusume_rank_defs():
    _out = []
    for _rk, _kd, _ax, _ats, _roi, _hit, _n, _typ in _OSUSUME_RANK_RAW:
        _atoms = []
        for _m, _t in _ats:
            _col, _lb = _OSU_ATOM[_m]
            _atoms.append((_col, float(_t), f'{_lb}{_t:g}'))
        _out.append(dict(
            tag=f'おすすめ#{_rk}', kind=_kd, axis=_ax, atoms=_atoms,
            roi=float(_roi), hit=float(_hit), n=int(_n),
            src='osusume', rank=_rk, typ=_typ))
    return _out


_OSUSUME_RANK_DEFS = _build_osusume_rank_defs()   # ★v136_001: 未使用(下で全面差し替え)

# ══════════════════════════════════════════════════════════
# ★v136_001: 次点本線枠を『最適紐008 再分析』の提案5件へ全面差し替え。
#   旧・有望条件.json 由来の J コード群と おすすめ#1〜30(最適紐007)は
#   使用しない(ローダ・定義は残置。この代入を消せば従来動作に戻る)。
#
#   次点枠に置く基準 = 「的中率か回収率のどちらかが本線基準に届かない」
#                      または「期間バイアスの疑いが晴れていない」条件。
#   実弾ではなく表示専用(賭け金・財布を通らない)。1〜2か月フォワードで観察する。
#
#   軸コード: wa50 = 補正後 単勝確率50%以上 / wb40 = 補正後 単勝確率40〜49.9%
#             ti90 = 単指数90以上(オッズ足切りなし)
#   ★相手アトムは全て補正後の列(単勝確率/複勝確率/単指数/複指数/紐馬指数)。
#     ただし『騎手補正』のみ表示用の偏差値列(KISHU_DEV_COL)を使う。
#     最適紐008の「騎手指数」は予想表HTMLの『騎手補正』列を指しており、
#     本コードの内部列 '騎手指数'(生・部外秘)とは別物のため取り違えないこと。
# [tag, kind, axis, [(列名, しきい値, ラベル), ...], roi, hit, n, タイプ, メモ]
# ══════════════════════════════════════════════════════════
_JITEN_V136_RAW = [
    # 全条件中トップの的中率。ただし回収率が100%を割るため本線には上げない。
    # 同じ買い目の馬連(HB1)・馬単(HB2)は本線枠で実弾採用済み。
    ('次#1', 'wide', 'wa50',
     [('複勝確率', 47.6, '複勝47.6')], 98.8, 64.9, 114, '高的中・回収不足',
     'ワイドは的中最高だがROI98.8%のトリガミ。100%を超えるか要観察'),
    # 数字は最良(的中62.7%/回収120.3%)だが、騎手補正の閾値を上げても成績が
    # 伸びない(41.2→48で的中58.3%へ低下)。旧テンプレート期を落とす日付フィルタ
    # として効いている疑いがあるため、期間分割の検証が済むまで次点に置く。
    ('次#2', 'wide', 'wb40',
     [('単勝確率', 15.6, '勝率15.6'), ('単指数', 68.0, '単指数68'),
      (KISHU_DEV_COL, 41.2, '騎手補正41.2')], 120.3, 62.7, 118, '高的中・要検証',
     '騎手補正の単調性なし=期間バイアス疑い。HA1の上位互換候補'),
    # 紐馬指数が単調に効く良条件(39.8→117.5 / 50.5→117.8 / 60→124.5 / 70→136.7)。
    # ただし的中26.7%と低く、的中率重視の方針では本線に置けない。
    ('次#3', 'umaren', 'ti90',
     [('複勝確率', 47.6, '複勝47.6'), ('紐馬指数', 70.0, '紐馬70')],
     136.7, 26.7, 105, '高回収・低的中',
     '紐馬指数が単調。回収は最良だが的中率が本線基準に届かない'),
    # 同系条件が軒並み108〜118%で揃っており筋は良いが n=91 と薄い。
    ('次#4', 'wide', 'ti90',
     [('複勝確率', 47.6, '複勝47.6'), ('複指数', 22.0, '複指数22'),
      ('紐馬指数', 60.0, '紐馬60')], 114.7, 44.0, 91, 'バランス',
     'n=91と標本が薄い。近傍条件は108〜118%で安定'),
    # 馬連で46%は高いが回収がちょうど100%割れ。しきい値を細かく刻めば抜ける可能性。
    ('次#5', 'umaren', 'wa50',
     [('単勝確率', 15.6, '勝率15.6'), ('単指数', 58.0, '単指数58')],
     99.8, 46.3, 81, '高的中・回収不足',
     '馬連46.3%は高水準。ROI99.8%で惜しい。しきい値の細分化で再探索したい'),
]

JITEN_MAP_DEFS = [
    dict(tag=_t, kind=_k, axis=_ax, atoms=[(_c, float(_th), _lb) for _c, _th, _lb in _ats],
         roi=float(_roi), hit=float(_hit), n=int(_n), src='v136', rank=_r, typ=_ty, memo=_mm)
    for _r, (_t, _k, _ax, _ats, _roi, _hit, _n, _ty, _mm) in enumerate(_JITEN_V136_RAW, start=1)
]
# ★v136_018: 次点本線枠 = 実戦条件シートの『観察のみ』2条件(表示専用・実弾外)。
#   ここまでの JITEN_MAP_DEFS(有望条件.json / おすすめランキング / 最適紐008)は使わない。
#   ・axis 'S:...' は本線と同じ選び方の軸(_jiten_axis で解決)。
#   ・atoms の 'S:列名' は予想表の表示値(小数1桁)で >= 判定。'_odds_le' は表示の単オッズ(騎手補正後) <= 閾値。
#   ・top1_himo=True: 該当馬のうち紐馬指数1位の1頭だけにする(探索と同じ)。
JITEN_MAP_DEFS = [
    dict(tag='OW1', kind='wide', axis='S:ti90max_o19',
         atoms=[('S:複勝確率', 40.0, '複勝確率40以上'), ('_odds_le', 30.0, '単勝オッズ30倍以内')],
         top1_himo=True, src='sheet', typ='観察', roi=115.1, hit=64.6, n=82),
    dict(tag='OU1', kind='umaren', axis='S:aw50max_o19',
         atoms=[('S:単勝確率', 15.0, '単勝確率15以上'), ('S:紐馬指数', 60.0, '紐馬指数60以上')],
         top1_himo=True, src='sheet', typ='観察', roi=117.4, hit=52.5, n=61),
]
print(f'[次点本線枠] ★v136_018: 実戦条件シートの観察条件 {len(JITEN_MAP_DEFS)} 件(表示のみ)。')


def gz_tier(cname):
    return GZ_COND_DEFS.get(cname, {}).get('tier', 'cover')


GZ_MAIN_NAMES = [k for k, v in sorted(GZ_COND_DEFS.items(), key=lambda kv: kv[1]['order'])
                 if v['tier'] == 'main']
GZ_ANA_NAMES = [k for k, v in sorted(GZ_COND_DEFS.items(), key=lambda kv: kv[1]['order'])
                if v['tier'] != 'main']
# ★v135_004: Kelly撤去に伴い GZ_SHRINK_K は不要。レジストリの p/pt/lo/hit は
#   『表示・記録用の検証統計』として保持する(賭け金計算には一切使わない)。

GZ_UMATAN_NAMES = [k for k, v in sorted(GZ_COND_DEFS.items(), key=lambda kv: kv[1]['order'])
                   if v['kind'] == 'umatan']
GZ_UMAREN_NAMES = [k for k, v in sorted(GZ_COND_DEFS.items(), key=lambda kv: kv[1]['order'])
                   if v['kind'] == 'umaren']
# ★v136_001: 本線枠のワイド(HA1)。馬単/馬連と同じ経路で買い目を生成する。
GZ_WIDE_NAMES = [k for k, v in sorted(GZ_COND_DEFS.items(), key=lambda kv: kv[1]['order'])
                 if v['kind'] == 'wide']


def generate_analysis(group):
    if len(group) == 0:
        return {
            'table': pd.DataFrame(), 'axis_analysis': 0,
            'axis_class': 'D級軸（要注意）', 'payout_level_score': 0,
            'arare_class': '本命戦', 'race_trend': '',
            'special_single': 'なし', 'fuku_recs': [],
            'umatan_recs': [], 'umaren_recs': [], 'wide_recs': [],
            'tan_recs': [], 'hole_fuku_recs': [],
            'high_confidence': False, 'conf_score': 0.0,
            'maruren_1pt': None, 'wide_1pt': None,
            'popular_out_high_payout': False, 'pop_conf_score': 0.0,
            'pop_maruren_recs': [], 'pop_wide_recs': [],
            'myomi_axis_score': 10, 'hole_bans': [], 'fav_bans': [],
            'judgment_class': 'C軸', 'recommended_action': '基本的にパス',
            'expected_win_rate': '—', 'expected_place_rate': '—',
            'judgment_comment': 'データなし', 'judgment_score': 0.0, 'axis_zscore': 0.0,
            'axis_win_prob': 0.0, 'axis_place_prob': 0.0,
            'gz_umatan_cond': {}, 'gz_umaren_cond': {},
            'anaba_hidx_bans': [],
            'gz_stakes': {},
            'gz_active_conds': [], 'gz_bankroll': GZ_BANKROLL,
            'gz_flat_unit': GZ_FLAT_UNIT, 'gz_max_points': gz_race_max_points(),
            'gz_dropped_n': 0, 'gz_race_invest': 0,
            'gz_active_main': [], 'gz_active_ana': [],
            'gz_main_points': 0, 'gz_ana_points': 0,
            'gz_main_invest': 0, 'gz_ana_invest': 0,
            'gz_ana_max_points': gz_ana_max_points(),
        }

    group = group.reset_index(drop=True)
    # ★v136_021: 出走取消/除外(出馬表HTML)を除き、騎手変更(horselist)を反映してから計算する
    _dq_scratched = {}
    try:
        _dq_scratched = stat_scratched_for_group(group)
        if _dq_scratched and len(group) - len(_dq_scratched) >= 2:
            group = group[~group['番'].map(lambda b: _to_int_z2h(b) in _dq_scratched)].reset_index(drop=True)
        else:
            _dq_scratched = {}
    except Exception as _e_scr:
        print(f'  [出走取消] 判定に失敗: {_e_scr}')
        _dq_scratched = {}
    try:
        _dq_jk = _apply_jockey_changes(group)
    except Exception as _e_jk:
        print(f'  [騎手変更] 判定に失敗: {_e_jk}')
        _dq_jk = []
    try:
        _dq_label = '%s%sR' % (re.sub(r'\s', '', unicodedata.normalize('NFKC', str(group['場所'].iloc[0]))),
                               _race_no_of(group['レース'].iloc[0]))
    except Exception:
        _dq_label = ''
    # ★NaN防御(v104): 通常はCSV読込で0埋め済みだが、generate_analysis を直接呼ぶ場合に
    #   単指数等が NaN だと指数→BADO→複合スコア→各 int(round()) へ伝播して落ちるため、
    #   読込側(numeric_cols)と同じく主要数値列の NaN を 0 補完して安全化する(正常データでは無変化)。
    for _ncol in ('単指数', '複指数', '騎手指数', '単勝オッズ', '単勝人気', '番'):
        if _ncol in group.columns:
            group[_ncol] = pd.to_numeric(group[_ncol], errors='coerce').fillna(0)
    group['印'] = ''
    group['unfold_score'] = group.apply(lambda row: get_dynamic_unfold_score(row['展開'], row['距離']), axis=1)
    # ── 騎手指数によるオッズ補正(v015) ──
    #   旧版は adjustment_factor がレース平均÷平均≈1.0 で全馬同一係数となり、
    #   人気順位が一切動かず実質無補正だった。ここを馬ごとの相対補正へ改める。
    #   人気騎手(平均より騎手指数が高い)=係数<1 でオッズを下げ、
    #   非人気騎手(平均より低い)          =係数>1 でオッズを上げる。
    jc = group['騎手指数'].clip(lower=1.0)
    j_mean = jc.mean()
    if pd.isna(j_mean) or j_mean <= 0:
        j_mean = 1.0
    if getattr(Config, 'JOCKEY_ODDS_ADJ_ENABLE', False):
        strength = float(getattr(Config, 'JOCKEY_ODDS_ADJ_STRENGTH', 0.30))
        cmin = float(getattr(Config, 'JOCKEY_ODDS_ADJ_MIN', 0.70))
        cmax = float(getattr(Config, 'JOCKEY_ODDS_ADJ_MAX', 1.50))
        # 係数 = (騎手指数 / レース平均) ^ (-strength)
        #   指数が平均より高い → 比>1 → 負べきで係数<1 → オッズ低下
        #   指数が平均より低い → 比<1 → 係数>1 → オッズ上昇
        ratio = (jc / j_mean).clip(lower=0.01)
        jockey_factor = np.power(ratio, -strength).clip(lower=cmin, upper=cmax)
    else:
        jockey_factor = pd.Series(1.0, index=group.index)
    group['jockey_odds_factor'] = jockey_factor.round(3)
    # ★【バグ修正】単勝オッズは 1.0倍 が下限(元返しでも100円=1.0倍)。
    #   旧コードは np.clip(..., 0, 199) と下限が 0 だったため、騎手補正の
    #   係数(最小 JOCKEY_ODDS_ADJ_MIN=0.70)が掛かると
    #     例) 予測1.3倍 × 0.70 = 0.91 → 表示「0.9倍」
    #   という実在しないオッズが出ていた(EV・人気順にも波及)。下限を 1.0 にする。
    _raw_odds = pd.to_numeric(group['単勝オッズ'], errors='coerce')
    _has_odds = (_raw_odds > 0).fillna(False)
    group['adjusted_odds'] = np.where(
        _has_odds,
        (_raw_odds.fillna(0.0) * jockey_factor).clip(1.0, 199.0).round(1),
        0.0)
    # ★【バグ修正】オッズ欠損(0埋め)の馬は、昇順ソートで先頭に来てしまい
    #   『補正後1番人気』に化けていた。欠損は最下位人気として末尾へ回す。
    #   (ソートは従来どおり安定ソート=同オッズは元の並び順を維持)
    group['_odds_sort'] = np.where(_has_odds, group['adjusted_odds'], np.inf)
    group = group.sort_values(
        '_odds_sort', kind='mergesort').drop(columns=['_odds_sort']).reset_index(drop=True)
    group['adjusted_popularity'] = range(1, len(group) + 1)

    z_tanshisu = race_zscore(group['単指数'])
    z_unfold = race_zscore(group['unfold_score'])
    p_tanshisu = race_percentile(group['単指数'])
    p_fukushisu = race_percentile(group['複指数'])
    p_jockey = race_percentile(group['騎手指数'])
    p_inv_odds = race_percentile(1 / group['単勝オッズ'].clip(lower=0.1))
    pos_raw = 18 - group['番']
    p_pos = race_percentile(pos_raw)

    # ══════════════════════════════════════════════════════════════
    # 単勝確率・複勝確率: LightGBM学習モデル(keiba_prob_model)で算出。
    #   ・特徴量: 馬番/単指数/複指数/騎手指数/脚質(展開)/距離/斤量/性別/年齢
    #   ・人気/オッズはモデルに入れない(期待値計算用に温存)。
    #   ・従来の経験補正ブレンド(emp_single/emp_place)は廃止し、モデル生確率を
    #     従来スケール(単勝合計100% / 複勝5〜95%クリップ・target_fuku_sum正規化)へ
    #     整えて後段(単勝期待値/グレード判定)との互換を保つ。
    # ══════════════════════════════════════════════════════════════
    _pm = _get_prob_model()
    _dist_val = group['距離'].iloc[0] if '距離' in group.columns and len(group) else 0
    # ★ 検証(keiba_prob_model.annotate_race)と同一経路で確率を算出する。
    #   単勝 = win_model の生予測をレース内で正規化(isotonicは適用しない=raw比)。
    #   複勝 = place_model にisotonic較正を適用した"生の3着内率"(正規化なし)。
    #   軸/相手の条件判定はこの生確率で行い、検証時と基準を一致させる。
    import numpy as _np
    # ★ モデル対応: 別AI予測オッズ(単勝オッズ)・別AI予測人気(単勝人気)を
    #   'odds'/'popularity' として渡す。これを渡さないと _row 側で欠損(-1)扱いになり、
    #   学習時に使った別AI特徴量が推論で死んで予測精度が大きく劣化する。
    # ★ horselist(血統/成績/タイム等)を各馬に結合。学習37特徴量のうち
    #   horselist由来27特徴を推論でも生かす(未結合だと全馬-1で予測が壊れる)。
    _hl_list = _attach_hl_to_group(group, CURRENT_RACECARD_DIR)
    _dq_hl_n = len(group)                                   # ★v136_021 結合率
    _dq_hl_m = sum(1 for _x in _hl_list if _x)
    _rows_X = _np.array([_pm._row({
        'num': r['番'], 'tan_idx': r['単指数'],
        # 複指数は学習側(USE_FUKU_IDX=False)で不使用。渡しても無視されるが明示的に外す。
        '騎手指数': r.get('騎手指数', 0), '展開': r.get('展開'),
        '斤量': r.get('斤量', 0), '性齢': r.get('性齢'),
        'odds': (r.get('単勝オッズ') if '単勝オッズ' in group.columns else None),
        'popularity': (r.get('単勝人気') if '単勝人気' in group.columns else None),
        'hl': _hl_list[_pos] if _pos < len(_hl_list) else None,
    }, _dist_val) for _pos, (_, r) in enumerate(group.iterrows())], dtype=float)
    from keiba_prob_model import _apply_isotonic as _apply_iso
    _win_pred = _np.asarray(_pm.win_model.predict(_rows_X), dtype=float)   # raw(0-1)
    _place_pred = _np.asarray(_pm.place_model.predict(_rows_X), dtype=float)
    # 単勝: raw をレース内で正規化して 0-100(合計100%)
    _wsum0 = _win_pred.sum()
    _win_norm = (_win_pred / _wsum0 * 100.0) if _wsum0 > 0 else _np.full(len(_win_pred), 100.0/max(len(_win_pred),1))
    # 複勝(生): 各馬に isotonic 較正を適用した3着内率 0-100(正規化なし)
    _place_raw_calib = _np.array([_apply_iso(float(p), _pm.place_calib) * 100.0 for p in _place_pred])

    # 検証と同一基準の"生確率"。表示・判定とも複勝確率=生の3着内率に統一。
    group['単勝確率_生'] = _win_norm.round(2)
    group['複勝確率_生'] = _place_raw_calib.round(2)
    # ★採用条件のEV atom用: 生EV = (単勝確率_生/100) × 単勝オッズ
    #   検証(keiba_prob_model の ev)と同一定義。オッズ欠損は None(未合致・安全側)。
    if '単勝オッズ' in group.columns:
        group['EV_生'] = (group['単勝確率_生'] / 100.0
                          * pd.to_numeric(group['単勝オッズ'], errors='coerce')).round(3)
    else:
        group['EV_生'] = float('nan')

    _win_raw = _win_norm.copy()

    # ── 単勝(表示=判定): レース内で1着1頭 → 生確率を正規化し合計100%に ──
    _wsum = _win_raw.sum()
    if _wsum > 0:
        group['単勝確率'] = (_win_raw / _wsum * 100).round(1)
    else:
        group['単勝確率'] = np.round(100.0 / max(len(group), 1), 1)

    # ── 複勝(表示=判定): モデルの生の3着内率(isotonic較正済み)をそのまま採用 ──
    #   ★ 従来スケール(正規化+5〜95%クリップ)は廃止。軸条件判定(複勝60+)と
    #     表示を一致させるため、複勝確率=生の3着内率(0〜100)に統一する。
    group['複勝確率'] = _place_raw_calib.round(1)

    # ══════════════════════════════════════════════════════════════
    # ★v136_003: アンサンブル単勝確率(◎◯決定専用。予想表の列には出さない)
    #   ・単勝オッズ由来確率 = 1/単勝オッズ をレース内で正規化(合計100%)。
    #     オッズ欠損(0/未入力)の馬は0%扱い(正規化の分母には含めない)。
    #   ・モデル予測単勝確率 = 上で算出した表示用『単勝確率』(補正後)。
    #   ・按分比率はオッズ由来40% / モデル予測60%(ENSEMBLE_*_WEIGHTで調整可)。
    #   ・脚質(展開)が『追』の馬は ENSEMBLE_OIKOMI_PENALTY 倍に減点。
    # ══════════════════════════════════════════════════════════════
    _raw_odds_ens = pd.to_numeric(group['単勝オッズ'], errors='coerce') if '単勝オッズ' in group.columns else pd.Series(np.nan, index=group.index)
    _has_odds_ens = (_raw_odds_ens > 0).fillna(False)
    _inv_odds_ens = np.where(_has_odds_ens, 1.0 / _raw_odds_ens.clip(lower=0.1), 0.0)
    _inv_sum_ens = float(np.nansum(_inv_odds_ens))
    if _inv_sum_ens > 0:
        _odds_prob_ens = (_inv_odds_ens / _inv_sum_ens) * 100.0
    else:
        # オッズが1件も無いレース(全欠損)はモデル予測のみへフォールバック。
        _odds_prob_ens = group['単勝確率'].to_numpy(dtype=float)
    group['単勝確率_オッズ由来'] = pd.Series(_odds_prob_ens, index=group.index).round(2)

    _ensemble_raw = (group['単勝確率_オッズ由来'].to_numpy(dtype=float) * ENSEMBLE_ODDS_WEIGHT
                     + group['単勝確率'].to_numpy(dtype=float) * ENSEMBLE_MODEL_WEIGHT)
    _oikomi_mask = group['展開'].astype(str).str.contains('追', na=False).to_numpy() if '展開' in group.columns else np.zeros(len(group), dtype=bool)
    _ensemble_raw = np.where(_oikomi_mask, _ensemble_raw * ENSEMBLE_OIKOMI_PENALTY, _ensemble_raw)
    # ★予想表には出さない内部列。◎(1位)/○(2位)の決定にのみ使用する。
    group['アンサンブル単勝確率'] = pd.Series(_ensemble_raw, index=group.index).round(2)

    group['単勝期待値'] = (group['adjusted_odds'] * (group['単勝確率'] / 100) * 100).round(0).astype(int)
    group['mark_score'] = calculate_ics(group)
    group['新S軸指数'] = round(
        0.30 * p_tanshisu + 0.21 * p_jockey + 0.17 * p_fukushisu +
        0.12 * z_unfold * 8 + 0.12 * p_inv_odds + 0.08 * z_tanshisu * 10, 1).clip(0, 100)
    unfold_bonus = group['unfold_score'].copy()
    mask = group['展開'].str.contains('差|追', na=False)
    unfold_bonus = np.where(mask, unfold_bonus * 1.08, unfold_bonus)
    unfold_bonus = pd.Series(unfold_bonus, index=group.index)
    group['紐馬指数'] = round(
        0.36 * p_fukushisu + 0.21 * p_jockey +
        0.17 * race_percentile(unfold_bonus) +
        0.14 * p_tanshisu + 0.12 * p_inv_odds, 1).clip(0, 100)
    axis_analysis_raw = (
        0.26 * p_tanshisu + 0.16 * p_jockey + 0.15 * p_fukushisu +
        0.12 * group['unfold_score'] + 0.12 * p_inv_odds +
        0.10 * (z_tanshisu * 10) + 0.09 * group['単勝確率']
    )
    group['軸信頼スコア'] = round(axis_analysis_raw.clip(0, 100), 1)
    group['joint_synergy_bonus'] = group.apply(lambda r: get_joint_synergy_bonus(r['単指数'], r['複指数']), axis=1)
    group['hybrid_axis_score'] = (
        0.473 * group['mark_score'] + 0.300 * group['軸信頼スコア'] +
        0.091 * group['単勝確率'] + 0.091 * p_pos + 0.045 * group['joint_synergy_bonus']
    ).clip(0, 100).round(1)
    group['BADO指数'] = group['hybrid_axis_score']

    # =====================================================
    # ★ NCS（New Composite Score）軸馬・紐馬選定ロジック v1.0（仕様書全面採用）
    #   旧複合スコア(_axis_score/_nyuma_score)は全廃。
    #   印・人気・オッズを一切使わず、純粋指数のみでNCSを算出する。
    #       NCS = 0.28*単指数 + 0.32*BADO指数 + 0.30*複勝率_予測 + 0.10*勝率
    #   （複勝率_予測=複勝確率, 勝率=単勝確率）
    #   欠損値はレース平均で補完（レース平均が無ければ0）。
    # =====================================================

    def _fill_mean(col):
        s = pd.to_numeric(group[col], errors='coerce')
        m = s.mean()
        if pd.isna(m):
            m = 0.0
        return s.fillna(m).astype(float)

    _tanI  = _fill_mean('単指数')     # 単指数
    _bado  = _fill_mean('BADO指数')   # BADO指数
    _fukuP = _fill_mean('複勝確率')   # 複勝率_予測
    _winP  = _fill_mean('単勝確率')   # 勝率(予測)

    group['NCS'] = (0.28*_tanI + 0.32*_bado + 0.30*_fukuP + 0.10*_winP).round(1)
    # NCS順位（タイは馬番の小さい方を上位）
    _ncs_order = group.sort_values(['NCS', '番'], ascending=[False, True]).index.tolist()
    group['NCS順位'] = 0
    for _r, _idx in enumerate(_ncs_order, start=1):
        group.loc[_idx, 'NCS順位'] = _r

    # ===== v095: 軸馬 = 複合スコア1位(単指数*0.7 + BADO指数*0.3) に統一 =====
    #   軸選定・◎・買い目・軸馬サマリー・レース判定(級/gap/荒れ度)をすべて複合スコア基準に統一。
    #   NCSは各馬の表示列・紐ランキング素材としてのみ残す(軸選定には用いない)。
    group['複合スコア'] = (group['単指数'].astype(float) * 0.7 + group['BADO指数'].astype(float) * 0.3).round(1)
    _cs_order = group.sort_values(['複合スコア', '番'], ascending=[False, True]).index.tolist()
    axis_main_idx = _cs_order[0]
    _axis_set = {axis_main_idx}
    # 以降 axis_ncs は『軸の複合スコア』を表す(grade/gap/race_cat の判定基準)。
    axis_ncs = float(group.loc[axis_main_idx, '複合スコア'])
    _ncs2   = float(group.loc[_cs_order[1], '複合スコア']) if len(_cs_order) > 1 else 0.0
    _ncs3   = float(group.loc[_cs_order[2], '複合スコア']) if len(_cs_order) > 2 else 0.0
    ncs_gap = round(axis_ncs - _ncs2, 1)        # 複合スコア1位と2位の差
    gap_2to3 = round(_ncs2 - _ncs3, 1)          # 複合スコア2位と3位の差

    # ===== 2頭軸モード判定（v091） =====
    #   gap小でも「上位2頭がともに高水準 かつ 3位を明確に引き離す」場合は
    #   割れる見送りレースではなく『2頭軸で勝負する連系レース』とみなす。
    #   旧ロジックは ncs_gap<3 を一律D級・見送りに落としていたが、
    #   1位拮抗=実力上位2頭という構図を取りこぼしていたため分離する。
    two_axis_mode = (ncs_gap < 3 and axis_ncs >= 58.0
                     and _ncs2 >= 56.0 and gap_2to3 >= 4.0)
    # gap連続減点係数: gapが小さいほど軸信頼を滑らかに割り引く。
    #   gap>=6で減点なし(1.0)、gap=0で下限。2頭軸モードでは割引を緩和(下限0.85)。
    _gap_floor = 0.85 if two_axis_mode else 0.70
    gap_conf_factor = round(float(np.clip(_gap_floor + (1.0 - _gap_floor) * (ncs_gap / 6.0),
                                          _gap_floor, 1.0)), 3)

    # ===== 紐馬グループ（軸を除くNCS順位ベース） =====
    _non_axis_order = [i for i in _ncs_order if i not in _axis_set]
    himo_top    = _non_axis_order[:2]           # 最優先紐: NCS順位2位・3位
    himo_sub    = _non_axis_order[3:5]          # サブ紐:   NCS順位5位・6位
    himo_strong = []                            # 強力紐:   NCS4位 / NCS>=軸NCS*0.82 & BADO>=48
    if len(_non_axis_order) >= 3:
        himo_strong.append(_non_axis_order[2])  # NCS順位4位
    for _i in _non_axis_order:
        if _i in himo_top or _i in himo_strong:
            continue
        if float(group.loc[_i, 'NCS']) >= axis_ncs * 0.82 and float(group.loc[_i, 'BADO指数']) >= 48:
            himo_strong.append(_i)
    # 連系の標準紐(最優先+強力+サブ)を最大4頭まで保持
    himo_idx_list = []
    for _i in (himo_top + himo_strong + himo_sub):
        if _i not in himo_idx_list:
            himo_idx_list.append(_i)
    himo_idx_list = himo_idx_list[:5]

    # ===== 穴馬軸推奨判定 → 穴紐 =====
    hole_cands = detect_hole_candidates(group, _axis_set)
    is_hole_axis = bool(hole_cands)
    # 穴紐: 穴馬軸推奨=YES かつ NCS順位<=7、または NCS>=軸NCS*0.75
    ana_himo = []
    for _i in _non_axis_order:
        _ncs_i  = float(group.loc[_i, 'NCS'])
        _rank_i = int(group.loc[_i, 'NCS順位'])
        if (is_hole_axis and _rank_i <= 7) or (_ncs_i >= axis_ncs * 0.75):
            ana_himo.append(_i)
    ana_himo = ana_himo[:5]

    # ===== 荒れ度(配当水準) =====
    prob_std = group['単勝確率'].std(ddof=0) if len(group) > 1 else 5.0
    axis_pop = int(group.loc[axis_main_idx, 'adjusted_popularity'])
    payout_level_score = round(float(np.clip(
        20 + axis_pop * 5 + max(0.0, 8.0 - float(prob_std)) * 2.0, 5, 100)), 1)
    if payout_level_score <= 45:
        arare_class = '低配当予想（本命固め）'
    elif payout_level_score <= 60:
        arare_class = '中配当予想（やや荒れ）'
    else:
        arare_class = '高配当予想（大荒れ）'
    # ===== 軸グレード判定（v109: バックテスト最適化式 axis_grade_formula を採用） =====
    #   旧 _base_grade(複合スコア+gap閾値)+2頭軸/荒れ度補正は実成績と非単調(B>A等)だった
    #   ため廃止。S>A>B>C>D が3着内率で単調降順になるよう最適化した標準化線形スコア式で
    #   軸馬(複合スコア1位)の区分を決定する。特徴量は軸馬のこの時点の算出値を使用。
    ncs_grade = compute_axis_grade_letter(
        複合スコア=float(axis_ncs),
        gap=float(ncs_gap),
        複勝確率=float(group.loc[axis_main_idx, '複勝確率']),
        単勝確率=float(group.loc[axis_main_idx, '単勝確率']),
        単勝期待値=float(group.loc[axis_main_idx, '単勝期待値']),
        紐馬指数=float(group.loc[axis_main_idx, '紐馬指数']),
    )

    # ===== NCS上位集計 / レース分類（仕様書セクション4） =====
    _top3 = _ncs_order[:5]
    top3_avg = round(float(np.mean([float(group.loc[i, 'NCS']) for i in _top3])), 1) if _top3 else 0.0
    _top4 = _ncs_order[:5]
    cluster3 = sum(1 for i in _top4 if float(group.loc[i, 'NCS']) >= 55) >= 3

    if two_axis_mode and ncs_grade != 'D' and payout_level_score < 40 and top3_avg >= 48:
        race_cat = '有力レース（2頭軸）'
        recommended_action = '上位2頭軸で連系（馬連・ワイド）中心。単勝は見送り'
    elif ncs_grade == 'D' or ncs_gap < 3 or payout_level_score >= 40 or top3_avg < 48:
        race_cat = '見送りレース'; recommended_action = '基本的に見送り推奨'
    elif ncs_grade in ('S', 'A') and ncs_gap >= 7 and payout_level_score <= 30 and top3_avg >= 55:
        race_cat = '投資レース'; recommended_action = '期待回収率100%以上を狙う投資レース'
    elif (ncs_grade == 'A') or (ncs_gap >= 5 and payout_level_score <= 32) or (ncs_grade == 'B' and cluster3):
        race_cat = '有力レース'; recommended_action = '積極的に勝負（馬連・馬単中心）'
    else:
        race_cat = '標準レース'; recommended_action = 'ワイド中心に薄く'
    # ★表示用マスク(見送りレースの表示のみ一旦中止)
    race_cat, recommended_action = _mask_race_cat(race_cat, recommended_action)
    _spec_axis_type = 'NCS%s級軸' % ncs_grade
    judgment_class = '%s級軸 / %s' % (ncs_grade, race_cat)

    # ══════════════════════════════════════════════════════════════
    # ★買い目ロジック: 軸=採用条件ごとの軸atom該当馬(単勝確率_生 最大)。
    #   予想印は単勝確率_生 降順で ◎○▲△☆。単勝/複勝は安定運用のみ。
    # ══════════════════════════════════════════════════════════════

    # ── 単勝確率降順の並び(相手選定・印の基準) ──
    #   ※判定は検証と同一基準の"生確率"列を使う。並び順は生単勝(=表示単勝と同値)。
    _wp_order = group.sort_values(['単勝確率_生', '番'], ascending=[False, True]).index.tolist()

    # ══════════════════════════════════════════════════════════
    # ★v135_003: 採用条件(馬単6/馬連3) の軸決定
    #   軸ルール(JSONと同一): 軸atom該当馬のうち 単勝確率_生 が最大の馬。
    #   該当馬が無ければその条件は不発火(下位馬へ繰り上げない)。
    #   軸atomは4種類しか使わない:
    #     win_prob>=40      → U1 / U4
    #     win_prob>=50      → U2 / U3 / U5
    #     kishu_idx>=50     → U6 / R3
    #     EV_生 1.5以上     → R1 / R2
    # ══════════════════════════════════════════════════════════
    def _h_val_g(_i, col):
        """生列を float で取得(NaN/欠損は None)。_h_val の前段版。"""
        try:
            _v = float(group.loc[_i, col])
            return _v if _v == _v else None
        except Exception:
            return None

    def _axis_by_col(col, thr):
        """col(生列) >= thr を満たす馬のうち 単勝確率_生 最大のもの。"""
        for _i in _wp_order:                       # _wp_order は単勝確率_生 降順
            try:
                _v = float(group.loc[_i, col])
            except Exception:
                continue
            if _v == _v and _v >= thr:
                return _i
        return None

    def _axis_by_winprob(thr):
        return _axis_by_col('単勝確率_生', thr)

    # ★v136_001: 最適紐008 の軸は『補正後 単勝確率』で選ぶ。
    #   _axis_by_col は _wp_order(=単勝確率_生 降順)の先頭ヒットを返すため、
    #   補正後の列で最大の馬とは一致しないことがある。探索と定義を厳密に
    #   合わせるため、補正後 単勝確率が最大の1頭を直接選ぶ専用ヘルパを置く。
    def _adj_win(_i):
        try:
            _v = float(group.loc[_i, '単勝確率'])
        except Exception:
            return None
        return _v if _v == _v else None

    def _axis_adj_win(thr):
        """補正後 単勝確率 >= thr の馬のうち、その値が最大の1頭。同値は馬番昇順。"""
        _cand = [(_v, int(group.loc[_i, '番']), _i) for _i in group.index
                 for _v in [_adj_win(_i)] if _v is not None and _v >= thr]
        if not _cand:
            return None
        return min(_cand, key=lambda t: (-t[0], t[1]))[2]

    def _axis_adj_win_band(lo, hi):
        """補正後 単勝確率が [lo, hi) の帯にある馬のうち、その値が最大の1頭。
        hi 以上の馬がいるレースでも帯内に馬がいれば発火する点に注意
        (最適紐008 の『単勝確率40〜49.9%』軸と同一定義。50%以上の軸とは
         同じレースで併発しうるが、選ばれる軸馬は必ず別の馬になる)。"""
        _cand = [(_v, int(group.loc[_i, '番']), _i) for _i in group.index
                 for _v in [_adj_win(_i)] if _v is not None and lo <= _v < hi]
        if not _cand:
            return None
        return min(_cand, key=lambda t: (-t[0], t[1]))[2]

    # ★v136_005: 紐条件アーカイブ(最適紐016〜026)の軸。単指数(単一列・補正概念なし)を
    #   起点に、必要なら人気/オッズで絞り込む。選び方は _axis_adj_win と同じ原理――
    #   条件を満たす馬のうち『閾値をかけている指標(単指数)自身が最大』の1頭を選ぶ
    #   (_axis_by_col のように単勝確率_生 の順で選ぶわけではない点に注意)。
    #   人気・オッズは単勝人気/単勝オッズ(生の列。補正の概念がない)で判定する。
    #   ★v136_006: adj_win_min を追加(補正後 単勝確率の下限。V7 の複合軸用)。
    def _axis_tan_idx(thr, pop_max=None, odds_max=None, adj_win_min=None):
        _cand = []
        for _i in group.index:
            _tn = _h_val_g(_i, '単指数')
            if _tn is None or _tn < thr:
                continue
            if adj_win_min is not None:
                _aw = _adj_win(_i)
                if _aw is None or _aw < adj_win_min:
                    continue
            if pop_max is not None:
                _pp = (safe_pop(group.loc[_i, '単勝人気'])
                       if '単勝人気' in group.columns else float('nan'))
                if not (_pp == _pp and _pp <= pop_max):
                    continue
            if odds_max is not None:
                _od = _h_val_g(_i, '単勝オッズ')
                if _od is None or _od > odds_max:
                    continue
            _cand.append((_tn, int(group.loc[_i, '番']), _i))
        if not _cand:
            return None
        return min(_cand, key=lambda t: (-t[0], t[1]))[2]

    def _axis_by_ev(lo, hi=99.0):
        """EV_生 が [lo, hi) の馬のうち 単勝確率_生 最大のもの(band=lo<=v<hi)。
        ※_h_ev は後段定義のため、ここでは EV_生 を直接読む。"""
        for _i in _wp_order:
            try:
                _v = float(group.loc[_i, 'EV_生'])
            except Exception:
                continue
            if _v == _v and lo <= _v < hi:
                return _i
        return None

    _gz_axis_w40 = _axis_by_winprob(40.0)          # U1 の軸 / A50(継承)
    _gz_axis_w50 = _axis_by_winprob(50.0)          # U2 / U3 / A50 の軸
    _gz_axis_ks50 = _axis_by_col('騎手指数', 50.0)  # R3 の軸
    _gz_axis_ks60 = _axis_by_col('騎手指数', 60.0)  # ★v135_029: PC(提案③)の軸(補正前騎手指数)
    _gz_axis_ev15 = _axis_by_ev(1.5)               # R1 の軸
    # ★v135_014: 紐カバー枠を単勝確率軸テンプレへ全面変更したため軸atomを追加。
    #   軸ルールは既存と同一(該当馬のうち 単勝確率_生 最大の1頭)。
    # ★v135_023: A1 の軸。単指数90以上かつ予測オッズ1.5倍以下の馬のうち
    #   単勝確率_生 が最大の1頭。軸ルールは他と同一。
    #   ★単勝オッズ列はAIの予測オッズであり市場オッズではない。
    #     したがってこの軸は「市場の支持」ではなくモデル出力の合意を見ている。
    def _axis_tan90_odds15():
        # ★v135_031: 検証(hc8_fdr)と一致。単指数90+のうち単勝確率_生 最大の1頭を選び、
        #   その軸オッズが1.5を超えたら見送り(下位馬へ繰り上げない)。
        _ax = _axis_by_col('単指数', 90.0)
        if _ax is None:
            return None
        _o = _h_val_g(_ax, '単勝オッズ')
        if _o is None or _o > 1.5:
            return None
        return _ax

    # ★v136_001: 本線枠(HA1/HA2/HB1/HB2)の軸。
    _gz_axis_aw50 = _axis_adj_win(50.0)              # HB1 / HB2 の軸
    _gz_axis_aw40_49 = _axis_adj_win_band(40.0, 50.0)  # HA1 / HA2 の軸
    _gz_axis_a1 = _axis_tan90_odds15()
    _gz_axis_w20 = _axis_by_winprob(20.0)          # A20 の軸
    _gz_axis_w25 = _axis_by_winprob(25.0)          # A25 の軸
    _gz_axis_w30 = _axis_by_winprob(30.0)          # A30 の軸
    # ★v136_005: 本線枠(AR1〜AR9・紐条件アーカイブ)の軸。
    _gz_axis_ti90 = _axis_tan_idx(90.0)                        # AR2 の軸
    _gz_axis_ti90_pop1 = _axis_tan_idx(90.0, pop_max=1)        # AR5/AR6/AR7 の軸
    _gz_axis_ti90_pop1_o3 = _axis_tan_idx(90.0, pop_max=1, odds_max=3.0)  # AR8/V2/V3/V5 の軸
    # ★v136_006: 紐条件アーカイブ第2版で追加した軸。
    _gz_axis_ti90_pop1_o2 = _axis_tan_idx(90.0, pop_max=1, odds_max=2.0)  # V4/V6 の軸
    _gz_axis_aw50_ti90_pop1_o3 = _axis_tan_idx(90.0, pop_max=1, odds_max=3.0,
                                               adj_win_min=50.0)          # V7 の軸
    # ★v136_018: 実戦条件シートの軸。探索と同じ選び方で、軸を先に1頭決めてから条件を見る
    #   (満たさなければそのレースは見送り。下位馬へ繰り上げない)。
    #   ・単指数軸 = 単指数90以上のうち単指数最大(同値は馬番昇順) … _axis_tan_idx(90)
    #   ・単勝確率軸 = 表示の単勝確率50.0以上のうち最大(同値は馬番昇順) … _sheet_axis_aw50()
    #   ・判定は予想表に表示される値で行う(探索と同じ)。人気/オッズは騎手補正後
    #     (予想表の『人気』『単オッズ』列)。0/欠損は不成立。1.9倍以下 = 1.95未満。
    def _sheet_r1(_i, col):
        """予想表の表示と同じ小数1桁に丸めた値。"""
        _v = _h_val_g(_i, col)
        return None if _v is None else float('%.1f' % _v)

    def _sheet_ge(_i, col, thr):
        _v = _sheet_r1(_i, col)
        return _v is not None and _v >= thr

    def _sheet_odds_lt(_i, _lim):
        _o = _h_val_g(_i, 'adjusted_odds')
        return _o is not None and 0 < _o < _lim

    def _sheet_pop(_i):
        _p = _h_val_g(_i, 'adjusted_popularity')
        return 99.0 if _p is None else _p

    def _sheet_axis_aw50():
        """表示の単勝確率が50.0以上の馬のうち単勝確率最大(同値は馬番昇順)。"""
        _cand = [(_h_val_g(_i, '単勝確率'), int(group.loc[_i, '番']), _i)
                 for _i in group.index if _sheet_ge(_i, '単勝確率', 50.0)]
        if not _cand:
            return None
        return min(_cand, key=lambda t: (-t[0], t[1]))[2]

    _sheet_ti_ax = _axis_tan_idx(90.0)
    _sheet_aw_ax = _sheet_axis_aw50()
    _gz_axis_s_ti90_o19 = (_sheet_ti_ax if _sheet_ti_ax is not None
                           and _sheet_odds_lt(_sheet_ti_ax, 1.95) else None)          # UR1 / OW1
    _gz_axis_s_ti9093_o14 = (_sheet_ti_ax if _sheet_ti_ax is not None
                             and (_h_val_g(_sheet_ti_ax, '単指数') or 0) <= 93
                             and _sheet_odds_lt(_sheet_ti_ax, 1.45) else None)        # UR2
    _gz_axis_s_aw50_ti90_pop1 = (_sheet_aw_ax if _sheet_aw_ax is not None
                                 and (_h_val_g(_sheet_aw_ax, '単指数') or 0) >= 90
                                 and _sheet_pop(_sheet_aw_ax) == 1 else None)           # WD1
    _gz_axis_s_aw50_o19 = (_sheet_aw_ax if _sheet_aw_ax is not None
                           and _sheet_odds_lt(_sheet_aw_ax, 1.95) else None)          # OU1

    # 条件名 → 軸index の対応表(発火判定・買い目生成で参照)
    # ★v135_015: 本線枠・umatanカバー枠の軸atomは win_prob>=40/50 の2種のみ。
    #   本線枠とumatanカバー枠で軸を揃えることが紐カバー設計の前提条件。
    # ★v135_021: 馬連 R3 のために kishu_idx>=50 軸を再追加(_gz_axis_ks50)。
    #   ※これは umaren/cover 専用の例外。umatan 本線・カバー枠の軸atomを
    #     kishu 系に緩めてはならない(被覆率設計が壊れる)。
    _GZ_AXIS_IDX = {
        'win_prob>=40': _gz_axis_w40,
        'win_prob>=50': _gz_axis_w50,
        'kishu_idx>=50': _gz_axis_ks50,   # ★v135_021: 馬連R3の軸(例外)
        'kishu_idx>=60': _gz_axis_ks60,   # ★v135_029: PC(提案③)の軸(補正前騎手指数)
        'tan90_odds15': _gz_axis_a1,      # ★v135_023: A1の軸
        # ★v136_001: 最適紐008 由来の本線枠4条件の軸(いずれも補正後 単勝確率)。
        #   ・'adj_win>=50'   … 補正後 単勝確率50%以上のうち最大の1頭
        #   ・'adj_win40_49'  … 補正後 単勝確率40〜49.9%のうち最大の1頭(50%+とは別の馬)
        'adj_win>=50': _gz_axis_aw50,
        'adj_win40_49': _gz_axis_aw40_49,
        # ★v136_005: 紐条件アーカイブ(最適紐016〜026)の軸。
        #   ・'tan_idx>=90'          … 単指数90以上のうち単指数が最大の1頭(人気・オッズ制約なし)
        #   ・'tan_idx>=90&pop<=1'   … 単指数90以上 かつ 単勝人気1以内(生)
        #   ・'tan_idx>=90&pop<=1&odds<=3' … 上記に加え単勝オッズ3以内(生)
        'tan_idx>=90': _gz_axis_ti90,
        'tan_idx>=90&pop<=1': _gz_axis_ti90_pop1,
        'tan_idx>=90&pop<=1&odds<=3': _gz_axis_ti90_pop1_o3,
        # ★v136_006: 紐条件アーカイブ第2版。
        #   ・'tan_idx>=90&pop<=1&odds<=2' … 単指数90以上 & 人気1以内 & 単勝オッズ2以内
        #   ・'adj_win>=50&tan_idx>=90&pop<=1&odds<=3' … 上記オッズ3版に補正後勝率50%以上を追加
        'tan_idx>=90&pop<=1&odds<=2': _gz_axis_ti90_pop1_o2,
        'adj_win>=50&tan_idx>=90&pop<=1&odds<=3': _gz_axis_aw50_ti90_pop1_o3,
        # ★v136_018: 実戦条件シート
        'S:ti90max&odds<1.95': _gz_axis_s_ti90_o19,
        'S:ti90-93max&odds<1.45': _gz_axis_s_ti9093_o14,
        'S:aw50max&ti>=90&pop1': _gz_axis_s_aw50_ti90_pop1,
    }

    # 後段互換: 買い目(本線枠等)の軸選定には使わない旧代表軸候補(フォールバック専用)。
    umatan_axis_idx = _gz_axis_w40
    umaren_axis_idxs = [a for a in (_gz_axis_w50, _gz_axis_ks50, _gz_axis_ev15)
                        if a is not None]

    # ══════════════════════════════════════════════════════════
    # ★v136_003: ◎◯の決定 ── アンサンブル単勝確率の降順で固定する。
    #   ・◎ = アンサンブル単勝確率1位 / ○ = アンサンブル単勝確率2位。
    #   ・アンサンブル単勝確率 = 単勝オッズ由来確率40% + モデル予測単勝確率60%
    #     (脚質『追』は減点)。予想表の列には出さない(◎○決定専用)。
    #   ・買い目(本線枠/次点本線枠/単勝/複勝)の軸atomは従来どおり
    #     単勝確率_生ベースの _GZ_AXIS_IDX / _wp_order で決まり、この◎○とは
    #     独立(◎表示と買い目の軸は必ずしも一致しない)。
    # ══════════════════════════════════════════════════════════
    _ens_order = group.sort_values(
        ['アンサンブル単勝確率', '番'], ascending=[False, True]).index.tolist()
    if _ens_order:
        axis_idx = _ens_order[0]
    elif _wp_order:
        axis_idx = _wp_order[0]
    elif umatan_axis_idx is not None:
        axis_idx = umatan_axis_idx
    elif umaren_axis_idxs:
        axis_idx = umaren_axis_idxs[0]
    else:
        axis_idx = axis_main_idx
    second_idx = next((_i for _i in _ens_order if _i != axis_idx), None)
    axis_ban_wp = int(group.loc[axis_idx, '番'])

    # ══════════════════════════════════════════════════════════
    # ★v136_003: 予想印 ── ◎=アンサンブル単勝確率1位 / ○=同2位。
    #   ▲△☆(最大3頭)は『複勝確率(補正後)HIMO_PLACE_PROB_MIN%以上 または
    #   紐馬指数HIMO_HIMOBA_IDX_MIN以上』の馬(◎○を除く)へ、紐馬指数降順で
    #   付与する。該当馬が0頭のときのみ紐馬指数上位3頭にフォールバックする。
    #   ★参考予想(◎から馬連・ワイド)はここで確定した対象馬をそのまま使う。
    # ══════════════════════════════════════════════════════════
    def _hv_mark(_i, _col):
        try:
            _v = float(group.loc[_i, _col])
            return _v if _v == _v else None
        except Exception:
            return None

    group['印'] = ''
    group.loc[axis_idx, '印'] = '◎'

    # ★v136_010: 新ロジック ── ○▲△ = 参考買い目の相手3頭(印と買い目を完全連動)。
    #   ◎は従来どおり(アンサンブル単勝確率1位)。相手は ref_partner_order のスコア順。
    #   ☆は使わない。確率・指数の列は一切変更しない(印列のみ)。
    _ref_new_partners = None
    if REF_NEW_LOGIC:
        try:
            _odds_col_ref = 'adjusted_odds' if 'adjusted_odds' in group.columns else '単勝オッズ'
            _ord_ref = ref_partner_order(group['単勝確率'].to_numpy(),
                                         group[_odds_col_ref].to_numpy(),
                                         group.index.get_loc(axis_idx))
            _ref_new_partners = [group.index[_k] for _k in _ord_ref[:REF_N_PARTNERS]]
        except Exception as _re:
            print(f'  [参考予想警告] 新ロジックの相手選定に失敗したため旧ロジックを使用: {_re}')
            _ref_new_partners = None

    if _ref_new_partners is not None:
        # second_idx(=○) を先頭相手に置き換える。second_idx はこの印と参考予想でのみ使用。
        second_idx = _ref_new_partners[0] if _ref_new_partners else None
        _himo_is_fallback = False
        _himo_all_idxs = list(_ref_new_partners[1:])     # ○の後に並べる相手(▲△)
        _himo_partner_idxs = list(_ref_new_partners[1:])
        for _mk, _i in zip(REF_PARTNER_MARKS, _ref_new_partners):
            group.loc[_i, '印'] = _mk
    else:
        if second_idx is not None:
            group.loc[second_idx, '印'] = '○'

        _mark_excl = {axis_idx} | ({second_idx} if second_idx is not None else set())
        _himo_cand = [_i for _i in group.index if _i not in _mark_excl
                      and ((_hv_mark(_i, '複勝確率') or -1.0) >= HIMO_PLACE_PROB_MIN
                           or (_hv_mark(_i, '紐馬指数') or -1.0) >= HIMO_HIMOBA_IDX_MIN)]
        _himo_is_fallback = False
        if not _himo_cand:
            # 該当馬0頭のときのみ、紐馬指数上位3頭へフォールバック。
            _himo_is_fallback = True
            _himo_cand = [_i for _i in group.index if _i not in _mark_excl]
        _himo_cand = sorted(
            _himo_cand,
            key=lambda _i: (-(_hv_mark(_i, '紐馬指数') or -1.0), int(group.loc[_i, '番'])))
        # ★v136_004: 参考予想(◎からの馬連・ワイド)は『複勝確率47.7%以上or紐馬指数60以上』
        #   の該当馬”全員”を対象にする(該当馬が4頭以上いても漏らさない)。
        # ★v136_004mod: 条件該当馬は全員に必ず予想印を付ける。
        #   ・▲(1位) / △(2位) / ☆(3頭目以降)。4頭以上でも☆を継続付与。
        #   ・フォールバック時(複勝/紐馬指数条件に0頭の場合)は◎◯以外の紐馬指数上位3頭のみ。
        # ★v136_010 不具合修正: フォールバック時に参考予想の相手が『◎○以外の全頭』に
        #   なっていた(印は3頭なのに参考は全頭流し)。コメントどおり上位3頭に限定する。
        _himo_all_idxs = _himo_cand[:3] if _himo_is_fallback else list(_himo_cand)
        # フォールバック時は上位3頭のみ印付与、通常時は全員に印付与
        _himo_partner_idxs = _himo_cand[:3] if _himo_is_fallback else list(_himo_cand)
        _sub_marks = ['▲', '△', '☆']
        for _rank, _i in enumerate(_himo_partner_idxs):
            _mk = _sub_marks[min(_rank, len(_sub_marks) - 1)]
            group.loc[_i, '印'] = _mk


    # ── ★v133_modified011: 『穴』印 ──
    #   条件: 複指数=25 かつ 補正前単勝人気4-9 かつ 騎手指数27-62 かつ 単指数<=82
    #   ・無印馬 → 印『穴』 / 既印馬(○▲△☆) → 印に『穴』を追記(例: △穴)
    #   ・◎は軸表記を優先して印はそのまま(行ハイライトで強調)
    #   ・HTML/GUI/Excelでオレンジ強調表示。anaba_hidx_bans として返却。
    anaba_hidx_bans = []
    for _i in group.index:
        try:
            _fk = float(group.loc[_i, '複指数'])
            _tp = (safe_pop(group.loc[_i, '単勝人気']) if '単勝人気' in group.columns
                   else safe_pop(group.loc[_i, 'adjusted_popularity']))
            _kj = float(group.loc[_i, '騎手指数'])
            _tn = float(group.loc[_i, '単指数'])
        except Exception:
            continue
        if (_fk == _fk and int(round(_fk)) == 25 and 4 <= _tp <= 9
                and 27 <= _kj <= 62 and _tn <= 82):
            anaba_hidx_bans.append(int(group.loc[_i, '番']))
            _cur = str(group.loc[_i, '印'] or '')
            if _cur == '':
                group.loc[_i, '印'] = '穴'
            elif _cur != '◎':
                group.loc[_i, '印'] = _cur + '穴'

    # ══════════════════════════════════════════════════════════
    # ★v135_003: 採用条件 の相手(紐)判定
    #   紐atomは検証(JSON)と完全に同一。band は lo<=v<hi。
    # ══════════════════════════════════════════════════════════
    def _h_pop(_i):
        # 補正後人気(adjusted_popularity)を採用。単勝人気があればそちら優先。
        # ★欠損(0埋め)は NaN。NaN <= n は False になるため、人気不明の馬が
        #   人気上限の紐条件をすり抜けるのを防ぐ。
        if '単勝人気' in group.columns:
            _p = safe_pop(group.loc[_i, '単勝人気'])
            if _p == _p:
                return _p
        return safe_pop(group.loc[_i, 'adjusted_popularity'])

    def _h_ev(_i):
        v = group.loc[_i, 'EV_生']
        try:
            v = float(v)
            return v if v == v else None   # NaN→None
        except Exception:
            return None

    def _h_val(_i, col):
        """生列を float で取得(NaN/欠損は None)。"""
        try:
            _v = float(group.loc[_i, col])
            return _v if _v == _v else None
        except Exception:
            return None

    def _h_odds(_i):
        """補正前(元)単勝オッズ。

        ★v135_010: 以前は列が無いとき adjusted_odds(騎手補正済み)で代替して
          いたが、補正係数0.70-1.50のぶん帯[10,30)の判定が検証とズレる。
          EV_生 側は既に『列が無ければ不成立』なので挙動を揃え、列が無ければ
          None(=条件不成立・安全側)を返す。
        """
        return _h_val(_i, '単勝オッズ') if '単勝オッズ' in group.columns else None

    def _ge(_i, col, thr):
        _v = _h_val(_i, col)
        return _v is not None and _v >= thr

    def _odds_in(_i, lo, hi):
        _v = _h_odds(_i)
        return _v is not None and lo <= _v < hi

    def _ev_in(_i, lo, hi=99.0):
        _v = _h_ev(_i)
        return _v is not None and lo <= _v < hi

    def _pop_le(_i, thr):
        """★v135_011b: 素の市場人気(単勝人気) <= thr。
        JSONの pop atom は補正前の市場人気で検証されているため、
        騎手補正後の adjusted_popularity ではなく生の列を使う。
        欠損(0/非数値)は safe_pop が NaN を返すので自動的に不成立。"""
        if '単勝人気' not in group.columns:
            return False
        _v = safe_pop(group.loc[_i, '単勝人気'])
        return _v == _v and _v <= thr

    # ── 紐判定関数(条件名 → 相手indexリスト) ──
    #   いずれも軸馬自身を除き、単勝確率_生 降順(=_wp_order)で返す。
    def _partners_S2(_axis):   # 複勝確率_生 上位2頭（★パラメータなし）
        """★v135_012: 検証で導出した本線条件。閾値を一切持たない。
        軸(単勝確率1位かつ40%以上)を除き、複勝確率_生 の高い順に2頭。
        同値は馬番の小さい方を優先(検証時の tie-break と一致させる)。"""
        def _fp(_i):
            _v = _h_val(_i, '複勝確率_生')
            return _v if _v is not None else -1.0
        _cand = [_i for _i in group.index if _i != _axis]
        _cand.sort(key=lambda _i: (-_fp(_i), int(group.loc[_i, '番'])))
        return _cand[:2]

    def _partners_F4(_axis):   # 複勝確率_生上位2 ∩ 期待値上位2
        """★v135_013: S2の2頭から期待値の低い方を落とす。
        検証コード(himo_intersect_v13.py の F4_純EV一致)と厳密に一致させる:
          z_i  = レース内z-score( logit(複勝確率_生/100) )  ※軸を含む全馬
          ev_i = log(sigmoid(c*z_i + b0)) + slope*log(単勝オッズ_i)
        単勝オッズ欠損はレース中央値で補完(全欠損なら発火しない)。"""
        try:
            return _partners_F4_impl(_axis)
        except Exception as _e:
            # ★絶対にレースを落とさない。F4は「買わない」に倒して続行する。
            global _F4_ERR_WARNED
            if not _F4_ERR_WARNED:
                _F4_ERR_WARNED = True
                print(f'[F4] 計算に失敗したため F4 を無効化して続行します: '
                      f'{type(_e).__name__}: {_e}')
            return []

    def _partners_F4_impl(_axis):
        _P = _load_f4_params()
        if not _P:
            return []
        _s2 = _partners_S2(_axis)
        if not _s2:
            return []
        _idx = list(group.index)
        # --- logit(複勝確率) の z-score（軸を含む全馬） ---
        _lg = []
        for _i in _idx:
            _v = _h_val(_i, '複勝確率_生')
            _q = 0.0 if _v is None else float(_v) / 100.0
            _q = min(max(_q, 1e-4), 1.0 - 1e-4)
            _lg.append(_math.log(_q / (1.0 - _q)))
        _n = len(_lg)
        _m = sum(_lg) / _n
        _sd = (sum((_x - _m) ** 2 for _x in _lg) / _n) ** 0.5
        if _sd < 1e-9:
            return []
        # --- 単勝オッズ（欠損はレース中央値で補完） ---
        _od = {}
        for _i in _idx:
            _o = (safe_odds(group.loc[_i, '単勝オッズ'])
                  if '単勝オッズ' in group.columns else None)
            _od[_i] = (float(_o) if (_o is not None and _o == _o and _o > 0)
                       else None)
        _valid = sorted(_v for _v in _od.values() if _v is not None)
        if not _valid:
            return []
        _med = _valid[len(_valid) // 2]
        # --- 期待値スコア ---
        _ev = {}
        for _k, _i in enumerate(_idx):
            if _i == _axis:
                continue
            _o = _od[_i] if _od[_i] is not None else _med
            _u = _P['c'] * ((_lg[_k] - _m) / _sd) + _P['b0']
            _u = max(min(_u, 30.0), -30.0)
            _lp = -_math.log1p(_math.exp(-_u))          # = log(sigmoid(u))
            _ev[_i] = _lp + _P['slope'] * _math.log(max(_o, 1.01))
        if not _ev:
            return []
        _top2 = set(sorted(_ev, key=lambda _i: (-_ev[_i],
                                                int(group.loc[_i, '番'])))[:2])
        return [_i for _i in _s2 if _i in _top2]

    # ★_partners_U3 は v135_019 で削除(実質11点・S2と93.8%重複)。
    # ★v135_020: _partners_U8 / _partners_M1 は U1 に統合済みのため削除。
    #   U8(単70) ⊆ M1(単60) ⊆ U1(単50) の入れ子で、U1 が全部を含む。
    # ── ★v135_023: 新本線枠の相手 ──
    def _partners_A1(_axis):
        """複勝確率_生 上位2頭。軸を除いてから順位を取る。
        ★この順序は検証側 (_select_partners) と一致させること。
          先に順位を取ると軸自身が枠を潰して1点になる。"""
        def _fp(_i):
            _v = _h_val(_i, '複勝確率_生')
            return _v if _v is not None else -1.0
        _cand = [_i for _i in group.index if _i != _axis]
        _cand.sort(key=lambda _i: (-_fp(_i), int(group.loc[_i, '番'])))
        return _cand[:2]

    def _partners_W1(_axis):
        """単指数>=50 & 騎手指数>=50。ただし出走11頭未満は発火しない。
        ★頭数ゲートはレース単位の条件なので相手関数側で落とす。
          検証では 11-12頭122.6% / 9-10頭114.2% / ≤8頭67.1%。"""
        if len(group.index) < 11:
            return []
        return [_i for _i in _wp_order if _i != _axis
                and _ge(_i, '単指数', 50.0) and _ge(_i, '騎手指数', 50.0)]

    def _partners_U1(_axis):   # 単指数>=50 & 騎手指数>=50
        # U8/M1/U1 統合後。検証N=1153 / 本線枠で最も母数が厚い。
        return [_i for _i in _wp_order if _i != _axis
                and _ge(_i, '単指数', 50.0) and _ge(_i, '騎手指数', 50.0)]

    # ── ★v135_021: 本線枠に追加(axis_enhanced 探索・FDR再判定由来) ──
    #   軸は win_prob>=50(_GZ_AXIS_IDX で解決)。相手は閾値ベースで、
    #   S2/F4(複勝確率上位2頭・順位ベース)とは選択ロジックが異なり包含しない。
    def _partners_P1(_axis):   # 複勝確率_生>=40 & 単指数>=60
        # 検証ROI140.5% / N200 / 的中51(採用条件中で最良p=0.012)。
        return [_i for _i in _wp_order if _i != _axis
                and _ge(_i, '複勝確率_生', 40.0) and _ge(_i, '単指数', 60.0)]

    def _partners_P2(_axis):   # 複勝確率_生>=40 & 単勝人気<=7
        # 検証ROI133.6% / N236 / 的中59。相手を人気サイドに寄せた版。
        return [_i for _i in _wp_order if _i != _axis
                and _ge(_i, '複勝確率_生', 40.0) and _pop_le(_i, 7)]

    def _partners_R3(_axis):   # 単勝確率_生>=40 & 単勝オッズ 1-10 (馬連・軸=騎手50)
        # 検証ROI127.5% / N342 / 的中74。軸が騎手指数のため分散目的でカバー枠へ。
        return [_i for _i in _wp_order if _i != _axis
                and _ge(_i, '単勝確率_生', 40.0) and _odds_in(_i, 1.0, 10.0)]

    # ── ★v135_028: 本線枠(勝ち筋マップ集約)の相手 ──
    def _partners_MB1(_axis):   # 複勝確率_生>=40 のみ(馬単・軸wp50+)
        # J01/J03/J04/J05/J06/J07/J11 の和集合。place40 に追加フィルタを掛けた
        # 各条件は全て本集合の部分集合であり、これ1本で同じ買い目に集約される。
        return [_i for _i in _wp_order if _i != _axis
                and _ge(_i, '複勝確率_生', 40.0)]

    def _partners_MB2(_axis):   # 複勝確率_生>=40 & 騎手指数>=50(馬連・軸wp50+)
        # J16(検証ROI111.5 / 的中41.9%)。馬連の勝ち筋骨格。
        return [_i for _i in _wp_order if _i != _axis
                and _ge(_i, '複勝確率_生', 40.0) and _ge(_i, '騎手指数', 50.0)]

    # ── ★v135_029: 本線枠(骨格 提案①②③)の相手 ──
    #   ★騎手指数・単勝人気は補正前(生)で判定する:
    #       騎手指数 = _ge(_i,'騎手指数',..)  → group['騎手指数'](生・補正なし)
    #       人気     = _pop_le(_i,..)         → group['単勝人気'](生の市場人気)
    def _partners_PA(_axis):   # 騎手指数>=40 & 単勝確率_生>=15 (提案①)
        # 軸wp50+ の本命から、騎手指数40以上かつモデル勝率15%以上の実力馬へ流す。
        # 検証ROI123.2 / 的中27.8% / N250。骨格の代表(回収率と再現性のバランス最良)。
        return [_i for _i in _wp_order if _i != _axis
                and _ge(_i, '騎手指数', 40.0) and _ge(_i, '単勝確率_生', 15.0)]

    def _partners_PB(_axis):   # 単勝確率_生>=15 & 単勝人気<=2 (提案②)
        # 相手を補正前の市場人気2以内に絞り、点数を減らして的中率を上げる版。
        # 検証ROI121.2 / 的中35.1% / N185。
        return [_i for _i in _wp_order if _i != _axis
                and _ge(_i, '単勝確率_生', 15.0) and _pop_le(_i, 2)]

    def _partners_PC(_axis):   # 複勝確率_生>=20 & 単指数>=60 (提案③)
        # 軸=補正前騎手指数60以上の1頭から、複勝率20%以上かつ単指数60以上へ広く流す。
        # 検証ROI112.6 / 的中8.6% / N8846(本骨格で最大標本・再現性重視)。
        return [_i for _i in _wp_order if _i != _axis
                and _ge(_i, '複勝確率_生', 20.0) and _ge(_i, '単指数', 60.0)]

    # ── ★v136_001: 本線枠(最適紐008 提案)の相手 ──
    #   ★判定列は全て『補正後』(単勝確率/複勝確率/単指数)。'_生' 列ではない。
    #   ★騎手指数は使わない(期間バイアス疑い。ファイル冒頭の注記を参照)。
    def _partners_HA(_axis):
        """単勝確率(補正後)>=15.6 & 単指数>=68 — HA1(ワイド)/HA2(馬連) 共通。
        軸=補正後勝率40〜49.9%。閾値を1段ずつ緩めると的中率が
        54.0%→45.8%→43.9%→43.7% と単調に落ちる(=筋の通った条件)。
        平均紐点数は1.0点で、実質1点買いになる。"""
        return [_i for _i in _wp_order if _i != _axis
                and _ge(_i, '単勝確率', 15.6) and _ge(_i, '単指数', 68.0)]

    def _partners_HB(_axis):
        """複勝確率(補正後)>=47.6 — HB1(馬連)/HB2(馬単) 共通。
        軸=補正後勝率50%以上。単勝確率8.4や紐馬指数39.8を足しても
        対象114R→112R・成績ほぼ同値のため、複勝確率47.6の1項目だけで足りる。
        ワイドは的中64.9%と最高だが回収98.8%のため本線では買わない(次点#1で観察)。"""
        return [_i for _i in _wp_order if _i != _axis
                and _ge(_i, '複勝確率', 47.6)]

    # ── ★v136_005: 本線枠(紐条件アーカイブ・最適紐016〜026)の相手 ──
    #   ★判定列は全て『補正後』(単勝確率/複勝確率/単指数)。『騎手補正』のみ
    #     表示用の偏差値列(KISHU_DEV_COL)を使う(紐条件アーカイブの「騎手指数」に対応)。
    #   ★v136_005b: 紐条件アーカイブの実測値(pt/vroi/vn/vhit)は、原データの
    #     '上位N番手'列が全条件で1・紐馬ソート順=紐馬指数順(基準:上位1頭 N=1)
    #     で検証されている(原本xlsxで確認済み)。該当馬が複数いても『全員』を
    #     紐に流すと検証条件とズレるため、紐馬指数が最大の1頭だけに絞る。
    #     同値は馬番昇順(_axis_adj_win 等の既存tie-break規約と統一)。
    def _top1_himo(_cands):
        """該当馬候補から紐馬指数降順(同値は馬番昇順)で上位1頭だけを返す。"""
        if not _cands:
            return []
        _scored = [((_h_val_g(_i, '紐馬指数') if _h_val_g(_i, '紐馬指数') is not None
                     else -1.0), int(group.loc[_i, '番']), _i) for _i in _cands]
        return [min(_scored, key=lambda t: (-t[0], t[1]))[2]]

    def _partners_AR_place477(_axis):
        """複勝確率(補正後)>=47.7 — AR1(馬単)/AR4(馬連) 共通(◎)。
        該当馬のうち紐馬指数最大の1頭(紐条件アーカイブの基準N=1)。"""
        _cands = [_i for _i in _wp_order if _i != _axis
                  and _ge(_i, '複勝確率', 47.7)]
        return _top1_himo(_cands)

    def _partners_AR_place477_kishu481(_axis):
        """複勝確率(補正後)>=47.7 & 騎手補正>=48.1 — AR3(馬単・△要検証)。
        AR1に騎手補正48.1以上を追加すると回収率が上乗せされる新候補。
        対象R=77と小さく、紐条件アーカイブでも単一ファイル発見のため要検証。
        該当馬のうち紐馬指数最大の1頭(紐条件アーカイブの基準N=1)。"""
        _cands = [_i for _i in _wp_order if _i != _axis
                  and _ge(_i, '複勝確率', 47.7) and _ge(_i, KISHU_DEV_COL, 48.1)]
        return _top1_himo(_cands)

    def _partners_AR_tan156_tanidx68(_axis):
        """単勝確率(補正後)>=15.6 & 単指数>=68 — AR6(馬連)/AR7(ワイド・いずれも△要検証)。
        HA1/HA2と同一の紐条件だが、軸を単指数90以上&人気1以内に変更した組。
        該当馬のうち紐馬指数最大の1頭(紐条件アーカイブの基準N=1)。"""
        _cands = [_i for _i in _wp_order if _i != _axis
                  and _ge(_i, '単勝確率', 15.6) and _ge(_i, '単指数', 68.0)]
        return _top1_himo(_cands)

    def _partners_AR_tanidx68_kishu589(_axis):
        """単指数>=68 & 騎手補正>=58.9 — AR9(ワイド・△要検証)。対象R=40と小さい。
        該当馬のうち紐馬指数最大の1頭(紐条件アーカイブの基準N=1)。"""
        _cands = [_i for _i in _wp_order if _i != _axis
                  and _ge(_i, '単指数', 68.0) and _ge(_i, KISHU_DEV_COL, 58.9)]
        return _top1_himo(_cands)

    # ── ★v136_006: 本線枠(紐条件アーカイブ第2版)の相手 ──
    #   いずれも AR 系と同じく『条件該当馬のうち紐馬指数最大の1頭』(N=1)。
    #   判定列は補正後(単勝確率/複勝確率)・単指数・複指数(生の出馬表列)。
    #   騎手指数(騎手補正)を含む条件は v136_001 の方針(期間バイアス疑い)に従い採用しない。
    def _partners_V_tan84(_axis):
        """単勝確率(補正後)>=8.4 — V1(馬連◎)。"""
        return _top1_himo([_i for _i in _wp_order if _i != _axis
                           and _ge(_i, '単勝確率', 8.4)])

    def _partners_V_place477_ti50(_axis):
        """複勝確率(補正後)>=47.7 & 単指数>=50 — V2(馬連◎)。"""
        return _top1_himo([_i for _i in _wp_order if _i != _axis
                           and _ge(_i, '複勝確率', 47.7) and _ge(_i, '単指数', 50.0)])

    def _partners_V_tan84_ti68(_axis):
        """単勝確率(補正後)>=8.4 & 単指数>=68 — V3(馬連◎)。"""
        return _top1_himo([_i for _i in _wp_order if _i != _axis
                           and _ge(_i, '単勝確率', 8.4) and _ge(_i, '単指数', 68.0)])

    def _partners_V_tan46(_axis):
        """単勝確率(補正後)>=4.6 — V4(馬連◎)。"""
        return _top1_himo([_i for _i in _wp_order if _i != _axis
                           and _ge(_i, '単勝確率', 4.6)])

    def _partners_V_place477_ti68(_axis):
        """複勝確率(補正後)>=47.7 & 単指数>=68 — V5(馬連○)。"""
        return _top1_himo([_i for _i in _wp_order if _i != _axis
                           and _ge(_i, '複勝確率', 47.7) and _ge(_i, '単指数', 68.0)])

    def _partners_V_tan46_fuku20(_axis):
        """単勝確率(補正後)>=4.6 & 複指数>=20 — V7(ワイド○)。"""
        return _top1_himo([_i for _i in _wp_order if _i != _axis
                           and _ge(_i, '単勝確率', 4.6) and _ge(_i, '複指数', 20.0)])

    # ── ★v136_018: 実戦条件シートの相手(いずれも該当馬のうち紐馬指数1位の1頭) ──
    def _partners_S_UR1(_axis):
        """単勝確率(補正後)>=8 — UR1 馬連①。"""
        return _top1_himo([_i for _i in _wp_order if _i != _axis
                           and _sheet_ge(_i, '単勝確率', 8.0)])

    def _partners_S_UR2(_axis):
        """人気(騎手補正後)<=3 & 紐馬指数>=50 — UR2 馬連②。"""
        return _top1_himo([_i for _i in _wp_order if _i != _axis
                           and _sheet_pop(_i) <= 3 and _sheet_ge(_i, '紐馬指数', 50.0)])

    def _partners_S_WD1(_axis):
        """複勝確率(補正後)>=45 & 単指数>=65 — WD1 ワイド①。"""
        return _top1_himo([_i for _i in _wp_order if _i != _axis
                           and _sheet_ge(_i, '複勝確率', 45.0) and _sheet_ge(_i, '単指数', 65.0)])

    # ── ★v135_015: 紐カバー枠の相手判定 ──
    #   ★軸atomは全て win_prob>=40/50(本線枠と同一)。ここを緩めてはいけない。
    #   EV帯・オッズ帯は検証と同じ半開区間 lo<=v<hi。
    def _partners_E1(_axis):   # 単指数>=60 & EV_生 0-1
        # 軸が堅い(wp50+)ときは相手も人気側で決まるため低EV帯が有効。
        # 検証N=241 / 検証ROI111.5% / 的中21.2%(カバー枠で最高)。
        return [_i for _i in _wp_order if _i != _axis
                and _ge(_i, '単指数', 60.0) and _ev_in(_i, 0.0, 1.0)]

    def _partners_C1(_axis):   # 複勝確率_生>=20 & 人気<=5
        # 旧A50(複20&人気7)の人気を1段締めた版。的中16.9%→19.1%。
        return [_i for _i in _wp_order if _i != _axis
                and _ge(_i, '複勝確率_生', 20.0) and _pop_le(_i, 5)]

    def _partners_C2(_axis):   # 騎手指数>=40 & 人気<=5
        return [_i for _i in _wp_order if _i != _axis
                and _ge(_i, '騎手指数', 40.0) and _pop_le(_i, 5)]

    # ★_partners_C3 は v135_019 で削除(実質29点・U1と91.9%重複)。
    def _partners_E3(_axis):   # 複勝確率_生>=30 & EV_生 1.5+
        # 軸wp40+では相手が荒れる余地があるため高EV帯が有効(軸wp50+とは逆)。
        return [_i for _i in _wp_order if _i != _axis
                and _ge(_i, '複勝確率_生', 30.0) and _ev_in(_i, 1.5)]

    # ★_partners_E2 は v135_019 で削除(ROI 18.9% / CI上限53.2%)。

    _GZ_PARTNER_FN = {
        'A1': _partners_A1, 'W1': _partners_W1,   # ★v135_023
        'S2': _partners_S2, 'F4': _partners_F4,
        'U1': _partners_U1,
        'P1': _partners_P1, 'P2': _partners_P2,   # ★v135_021 本線枠追加
        'E1': _partners_E1, 'C1': _partners_C1, 'C2': _partners_C2,
        'E3': _partners_E3,
        'R3': _partners_R3,                         # ★v135_021 馬連カバー枠追加
        'MB1': _partners_MB1, 'MB2': _partners_MB2,  # ★v135_028 本線枠(勝ち筋マップ)
        'PA': _partners_PA, 'PB': _partners_PB, 'PC': _partners_PC,  # ★v135_029 本線枠(骨格①②③)
        # ★v136_001 本線枠(最適紐008): 券種違いは同じ相手関数を共有する。
        'HA1': _partners_HA, 'HA2': _partners_HA,
        'HB1': _partners_HB, 'HB2': _partners_HB,
        # ★v136_005 本線枠(紐条件アーカイブ・最適紐016〜026)。
        'AR1': _partners_AR_place477, 'AR4': _partners_AR_place477,
        'AR2': _partners_AR_place477,
        'AR3': _partners_AR_place477_kishu481,
        'AR5': _partners_AR_place477,
        'AR6': _partners_AR_tan156_tanidx68, 'AR7': _partners_AR_tan156_tanidx68,
        'AR8': _partners_AR_place477,
        'AR9': _partners_AR_tanidx68_kishu589,
        # ★v136_006 本線枠(紐条件アーカイブ第2版)。V6 は AR8 と同じ相手(複勝確率47.7)。
        'V1': _partners_V_tan84,
        'V2': _partners_V_place477_ti50,
        'V3': _partners_V_tan84_ti68,
        'V4': _partners_V_tan46,
        'V5': _partners_V_place477_ti68,
        'V6': _partners_AR_place477,
        'V7': _partners_V_tan46_fuku20,
        # ★v136_018 本線枠(実戦条件シート)
        'UR1': _partners_S_UR1,
        'UR2': _partners_S_UR2,
        'WD1': _partners_S_WD1,
    }
    # ★v135_022: 本線枠の縮小に合わせて実装側も絞る。
    #   下の assert を意味のあるチェックとして残すため、
    #   GZ_COND_DEFS と同じキー集合に揃えてから検査する。
    _GZ_PARTNER_FN = {_k: _v for _k, _v in _GZ_PARTNER_FN.items()
                      if _k in GZ_COND_DEFS}
    # ★整合チェック: 定義と実装の食い違いを起動時に検出する。
    assert set(GZ_COND_DEFS) == set(_GZ_PARTNER_FN), (
        'GZ_COND_DEFS と _GZ_PARTNER_FN の条件名が不一致: '
        f'{set(GZ_COND_DEFS) ^ set(_GZ_PARTNER_FN)}')

    # ===== 互換スコア(描画用) = NCS =====
    _comp_map = {int(group.loc[_i, '番']): float(np.clip(group.loc[_i, 'NCS'], 0, 100)) for _i in group.index}
    group['compatibility_score'] = group['番'].map(_comp_map)

    # ===== 表示テーブル =====
    #   v092x: 表示順を NCS順 → 複合スコア(単指数*0.7 + BADO指数*0.3)降順に変更。
    group['複合スコア'] = (group['単指数'].astype(float) * 0.7 + group['BADO指数'].astype(float) * 0.3).round(1)
    # ★v135_005: 表示専用の補正騎手指数(偏差値)。買い目条件では使わない。
    group[KISHU_DEV_COL] = calc_kishu_dev(group)
    # ★v136_019: 統計予想(出馬表HTMLを読み込んだときだけ値が入る)
    try:
        _stat = compute_stat_for_group(group)
    except Exception as _e_st:
        print(f'  [統計予想] 失敗: {_e_st}')
        _stat = None
    if _stat:
        group['統計勝率'] = group['番'].map(lambda b: _stat['p_stat'].get(int(b)) if pd.notna(b) else np.nan)
        group['融合勝率'] = group['番'].map(lambda b: _stat['p_fused'].get(int(b)) if pd.notna(b) else np.nan)
        group['統計印'] = group['番'].map(lambda b: _stat['marks'].get(int(b), '') if pd.notna(b) else '')
    else:
        group['統計勝率'] = np.nan
        group['融合勝率'] = np.nan
        group['統計印'] = ''
    # ★v136_020: 統計裏付け(本線/次点の軸の統計勝率) と 観察条件ST1〜ST6(統計1位を紐)
    _stat_ura = {}
    _stat_obs = []
    if _stat:
        try:
            _ps_all = {int(b): float(v) for b, v in _stat['p_stat'].items()}
            _ban_of = {_i: int(group.loc[_i, '番']) for _i in group.index if pd.notna(group.loc[_i, '番'])}
            _idx_of = {_b: _i for _i, _b in _ban_of.items()}
            _tp_sum = sum(float(_h_val_g(_i, '単勝確率') or 0.0) for _i in _ban_of) or 1.0
            for _ak, _ai in (('UR1', _gz_axis_s_ti90_o19), ('UR2', _gz_axis_s_ti9093_o14),
                             ('WD1', _gz_axis_s_aw50_ti90_pop1), ('OU1', _gz_axis_s_aw50_o19)):
                if _ai is None or _ai not in _ban_of:
                    continue
                _ab = _ban_of[_ai]
                if _ab not in _ps_all:
                    continue
                _stat_ura[_ak] = dict(ban=_ab, ps=_ps_all[_ab], ok=_ps_all[_ab] >= STAT_URA_MIN)
            for _od in STAT_OBS_DEFS:
                _ax = _stat_ura.get(_od['axis'])
                if not _ax:
                    continue
                _cands = [(_b, _p) for _b, _p in _ps_all.items() if _b != _ax['ban'] and _b in _idx_of]
                if not _cands:
                    continue
                _pb, _pps = min(_cands, key=lambda t: (-t[1], t[0]))   # 統計勝率1位(同値は馬番昇順)
                _pi = _idx_of[_pb]
                _tp = float(_h_val_g(_pi, '単勝確率') or 0.0)
                _ratio = (_pps / (100.0 * _tp / _tp_sum)) if _tp > 0 else None
                _ok = True
                if 'ax_min' in _od and _ax['ps'] < _od['ax_min']:
                    _ok = False
                if 'p_min' in _od and _pps < _od['p_min']:
                    _ok = False
                if 'fp_min' in _od and not _sheet_ge(_pi, '複勝確率', _od['fp_min']):
                    _ok = False
                if 'ratio_min' in _od and (_ratio is None or _ratio < _od['ratio_min']):
                    _ok = False
                _stat_obs.append(dict(tag=_od['tag'], kind=_od['kind'], axis=_od['axis'],
                                      axis_ban=_ax['ban'], ban=_pb, ok=_ok,
                                      p_ax=_ax['ps'], p_pt=_pps,
                                      ratio=(round(_ratio, 2) if _ratio is not None else None),
                                      fp=_sheet_r1(_pi, '複勝確率'),
                                      odds=_h_val_g(_pi, 'adjusted_odds'),
                                      label=_od['label'], n=_od['n'], hit=_od['hit'], roi=_od['roi']))
        except Exception as _e_ura:
            print(f'  [統計裏付け] 失敗: {_e_ura}')
        _ura_bans = {v['ban'] for v in _stat_ura.values() if v['ok']}
        if _ura_bans:
            group['統計印'] = [
                (str(_m or '') + STAT_URA_MARK) if (pd.notna(_b) and int(_b) in _ura_bans) else _m
                for _b, _m in zip(group['番'], group['統計印'])]
    table_columns = [
        '番', '馬名', '騎手', KISHU_DEV_COL, '展開', '単指数', '複指数',
        'BADO指数', 'アンサンブル単勝確率', '紐馬指数', '印', 'adjusted_popularity', 'adjusted_odds',
        '単勝確率', '複勝確率', '単勝期待値', 'hybrid_axis_score', 'NCS', 'NCS順位', '複合スコア',
        '統計勝率', '融合勝率', '統計印'
    ]
    # ★v133_modified: 予想表の並び順を単勝確率降順(同率は番順)に変更。
    table_df = group[table_columns].sort_values(['単勝確率', '番'], ascending=[False, True])
    table_df = table_df.rename(columns={'adjusted_popularity': '単勝人気', 'adjusted_odds': '単勝オッズ'})
    table_df['compatibility_score'] = table_df['番'].map(_comp_map)

    # ===== 軸馬サマリー(★v133_modified: ◎=単勝確率1位 基準に変更) =====
    axis_ban = int(group.loc[axis_idx, '番'])
    sub_axis_ban = None
    a_fukuP = float(group.loc[axis_idx, '複勝確率'])
    a_tanI  = float(group.loc[axis_idx, '単指数'])
    a_winP  = float(group.loc[axis_idx, '単勝確率'])       # ◎の単勝確率
    # 2位との単勝確率差(信頼度に使用)
    _wp_sorted = sorted([float(group.loc[_i, '単勝確率']) for _i in group.index], reverse=True)
    a_winP_gap = round(a_winP - (_wp_sorted[1] if len(_wp_sorted) > 1 else 0.0), 1)
    a_score = a_winP
    # 軸信頼度スコア = ◎の単勝確率(0-100)をそのまま採用
    axis_analysis = round(float(np.clip(a_winP, 0, 100)), 1)
    # 軸クラス: ★現行モデル(単勝確率1位軸)に整合する新式。
    #   単勝確率(勝ちやすさ) + 複勝確率(堅実さ) + 2位差(抜けている度)で判定。
    #   旧 axis_grade_formula(judgment_class由来)は単勝期待値の重みが最大の負(-0.79)で、
    #   EVが高い(=オッズ妙味ある)軸ほどD級に落ちる旧複合スコアモデル用の式。現行の
    #   単勝確率1位軸とは基準が食い違うため、EVを混ぜない素直な確率ベース判定に統一する。
    #   (EVは軸判定でなく買い目条件側で使う。)
    def _wp_grade_letter(wp, fp, gap):
        if wp >= 40 and fp >= 70 and gap >= 12: return 'S'
        if wp >= 32 and fp >= 60 and gap >= 8:  return 'A'
        if wp >= 24 and fp >= 50 and gap >= 4:  return 'B'
        if wp >= 17 and fp >= 40:               return 'C'
        return 'D'
    _wp_grade_desc = {'S': 'S級軸（1強・高信頼）', 'A': 'A級軸（有力）',
                      'B': 'B級軸（標準）', 'C': 'C級軸（やや不安）',
                      'D': 'D級軸（混戦・軸不適）'}
    _axis_grade_letter = _wp_grade_letter(a_winP, a_fukuP, a_winP_gap)
    # ★v136_008: 新判定(◎の推定勝率/3着内率)。軸級・レース分類・信頼度を ◎ 基準で統一。
    #   旧 race_cat は別の馬(複合スコア1位)を旧スケールの式で判定していたため置き換える。
    judge_est = None
    if USE_NEW_JUDGE:
        try:
            _ax_pos = group.index.get_loc(axis_idx)
            _odds_col = 'adjusted_odds' if 'adjusted_odds' in group.columns else '単勝オッズ'
            judge_est = compute_judge_estimates(group['単勝確率'].to_numpy(),
                                                group[_odds_col].to_numpy(), _ax_pos)
            _axis_grade_letter = judge_grade_letter(judge_est['est_win'])
            race_cat, recommended_action = judge_race_cat(
                judge_est['est_win'], judge_est['est_top3'],
                judge_est['quinella'], judge_est['rival_win'])
            race_cat, recommended_action = _mask_race_cat(race_cat, recommended_action)
        except Exception as _je:
            print(f'  [判定警告] 新判定の計算に失敗したため旧判定を使用: {_je}')
            judge_est = None
    axis_class = _wp_grade_desc[_axis_grade_letter]
    _g_std = group['単勝確率'].std(ddof=0)
    single_zscore = (a_winP - group['単勝確率'].mean()) / _g_std if _g_std and _g_std != 0 else 0.0
    myomi_axis_score = int(round(float(np.clip(a_winP_gap * 3.0 + (a_winP - 20.0) * 1.5, 0, 90))))

    # ★judgment_class(HTML/GUI/Excelの表示に使用)を新軸グレードに統一。
    #   ★v136_008: race_cat も新判定(◎基準)。USE_NEW_JUDGE=False なら従来の複合スコア系。
    judgment_class = '%s / %s' % (_wp_grade_desc[_axis_grade_letter], race_cat)

    _himo_bans_disp = [int(group.loc[_i, '番']) for _i in _wp_order[1:5]]
    race_trend = '◎[%d] 単勝確率%.1f%%(2位差%.1f) ＋ 相手%s。複勝確率%.0f%%/単指数%.0f' % (
        axis_ban, a_winP, a_winP_gap, _himo_bans_disp, a_fukuP, a_tanI)
    special_single = '◎%d 単勝確率%.1f%% 複勝確率%.0f%% 単指数%.0f' % (axis_ban, a_winP, a_fukuP, a_tanI)

    # ===== 本命強軸 判定（表示用フラグ） =====
    is_honmei_axis = bool(a_tanI >= 90 and a_winP >= 40)

    # ===== 連系生成ヘルパー _mk_pairs と 'hole_bans' 用 hole_axis_bans のみ算出 =====
    #   v101整理: 旧 v085/v092 系の買い目計算(単複ゲート / select_conn_partners 連系 /
    #   有力馬単 / 穴狙い anaba / 3連複)は、後段 v093 の選定で全面上書きされる死にコード
    #   だったため削除。買い目生成に必要な _mk_pairs と、返却 'hole_bans' に使う
    #   hole_axis_bans のみを残す。
    _dist = int(group['距離'].iloc[0]) if len(group) > 0 else 1600
    _head = len(group)

    # ★ v132.2: 予測オッズ(quinella/exacta)を _mk_pairs の表示オッズにも使うため先に定義。
    #   従来は表示オッズ=estimate_realistic_odds(単勝オッズの和ベース)、ゲート判定=quinella/exacta
    #   という二重基準で、表示だけが過大(例: 馬単exacta22倍 → 表示129倍)になっていた。
    #   表示もゲートと同じ予測オッズに統一し、食い違いを解消する。
    _gw_raw = {}
    for _i in group.index:
        _wp = float(group.loc[_i, '単勝確率']) or 0.0
        if _wp <= 0:
            _od0 = float(group.loc[_i, '単勝オッズ']) or 0.0
            _wp = (100.0 / _od0) if _od0 > 0 else 0.0
        _gw_raw[int(group.loc[_i, '番'])] = _wp
    _gw_tot = sum(_gw_raw.values())
    _gp = {b: (v / _gw_tot) for b, v in _gw_raw.items()} if _gw_tot > 0 else {b: 0.0 for b in _gw_raw}
    def _quinella_odds(a, b):
        pa = _gp.get(a, 0.0); pb = _gp.get(b, 0.0)
        if min(pa, pb) <= 0 or max(pa, pb) >= 1:
            return None
        q = pa * pb / (1 - pa) + pb * pa / (1 - pb)
        return (1.0 / q) if q > 0 else None
    def _exacta_odds(a, b):   # a→b(a が1着)
        pa = _gp.get(a, 0.0); pb = _gp.get(b, 0.0)
        if pa <= 0 or pa >= 1 or pb <= 0:
            return None
        e = pa * pb / (1 - pa)
        return (1.0 / e) if e > 0 else None

    def _mk_pairs(axis_i, partner_idx_list):
        _aod = float(group.loc[axis_i, 'adjusted_odds'])
        _afk = float(group.loc[axis_i, '複勝確率'])
        _apr = float(group.loc[axis_i, '単勝確率'])
        _abn = int(group.loc[axis_i, '番'])
        _ur = []; _ut = []; _wd = []
        for _i in partner_idx_list:
            if _i == axis_i:
                continue
            _t_ban = int(group.loc[_i, '番']); _t_odds = float(group.loc[_i, 'adjusted_odds'])
            _t_pop = int(group.loc[_i, 'adjusted_popularity'])
            _t_fuku = float(group.loc[_i, '複勝確率']); _t_prob = float(group.loc[_i, '単勝確率'])
            # ★ v132.2: 馬連=quinella予測 / 馬単=exacta予測 / ワイド=複勝確率ベース。
            #   予測が算出不能(端の確率)な場合のみ従来式にフォールバック。
            _o_wide = wide_odds_from_place(_afk, _t_fuku)
            _pq = _quinella_odds(_abn, _t_ban)
            _pe = _exacta_odds(_abn, _t_ban)
            if _pq is not None and _pq > 0:
                _o_um = round(_pq, 1)
            else:
                _o_um = estimate_realistic_odds(_aod, _t_odds, _t_pop, payout_level_score, _dist, _head, 'umaren')
            if _pe is not None and _pe > 0:
                _o_ut = round(_pe, 1)
            else:
                _o_ut = estimate_realistic_odds(_aod, _t_odds, _t_pop, payout_level_score, _dist, _head, 'umatan')
            _wh = round(_afk / 100.0 * _t_fuku / 100.0 * 100, 1)
            _uth = round(_apr / 100.0 * _t_prob / 100.0 * 100, 1)
            _urh = round(min(99.5, _uth * 2), 1)
            _ew = round(_wh / 100.0 * _o_wide * 100, 0)
            _eur = round(_urh / 100.0 * _o_um * 100, 0)
            _eut = round(_uth / 100.0 * _o_ut * 100, 0)
            _cp = float(group.loc[_i, 'NCS'])
            _ur.append(PairRec(_abn, _t_ban, _urh, _o_um, _eur, _cp, _eur))
            _ut.append(PairRec(_abn, _t_ban, _uth, _o_ut, _eut, _cp, _eut))
            _wd.append(PairRec(_abn, _t_ban, _wh, _o_wide, _ew, _cp, _ew))
        return _ur, _ut, _wd


    # ══════════════════════════════════════════════════════════════
    # ★ v133_modified 買い目生成(単勝確率ベース + 有力2条件)
    # ══════════════════════════════════════════════════════════════
    # 返却互換のための旧変数(未使用は空)
    tan_flags, fuku_flags = {}, {}
    yuryoku_umatan_pairs = []
    hole_axis_bans = set()
    nsl_tan_tier = ''
    nsl_fuku_hit = False

    def _single_of(_i):
        return SingleRec(ban=int(group.loc[_i, '番']),
                         hit_rate=round(float(group.loc[_i, '単勝確率']), 1),
                         odds=float(group.loc[_i, 'adjusted_odds']),
                         ev=int(round(float(group.loc[_i, '単勝期待値']))))

    # ── 単勝: 買い目は単勝確率1位の馬1点。 ──
    #   ★検証(win_prob>=40 & tan_idx>=90 は"単勝確率1位の馬"で測定)と一致させるため、
    #     単勝の対象・有力判定は◎(条件別軸)ではなく『単勝確率1位の馬』で行う。
    _tan_target_idx = _wp_order[0] if _wp_order else axis_idx   # 単勝確率1位
    _axis_tan_idx = float(group.loc[_tan_target_idx, '単指数'])   # 後段コメント生成で使用
    tan_recs = [_single_of(_tan_target_idx)]
    # ★v134: 安定運用判定用に軸馬(単勝確率1位)の生確率・補正前人気を取得。
    #   検証と同一の '_生' 列を使う(win_prob/place_prob)。人気=単勝人気(補正前)。
    try:
        _stable_axis_ban = int(group.loc[_tan_target_idx, '番'])
        _stable_win_prob = float(group.loc[_tan_target_idx, '単勝確率_生'])
        _stable_place_prob = float(group.loc[_tan_target_idx, '複勝確率_生'])
        # ★欠損(0埋め)を「1番人気」と誤認しないよう safe_pop で NaN 化 → None
        _sp = (safe_pop(group.loc[_tan_target_idx, '単勝人気'])
               if '単勝人気' in group.columns else float('nan'))
        _stable_raw_pop = int(_sp) if _sp == _sp else None
    except Exception:
        _stable_axis_ban = None
        _stable_win_prob = _stable_place_prob = None
        _stable_raw_pop = None
    # ★変更(要望): 単勝の「有力」判定は廃止。単勝は常に『参考』表示とする。
    #   (検証で単勝の有力条件はCI下限が100%を割り、統計的裏付けが無いため)
    yuuryoku_tan = set()
    # ★変更(要望): 単勝の有力判定を廃止したため、優先ban集合も空にする。
    #   (空にしないと 3079行付近で単勝が『有力』表示に戻ってしまう)
    tan_priority_bans = set()

    # 複勝は廃止
    fuku_recs = []
    yuuryoku_fuku = set()
    hole_fuku_recs = []

    # ══════════════════════════════════════════════════════════════
    # ★v135_003: 採用条件 の買い目生成
    #   馬単 = U1 ∪ U2 ∪ U3 ∪ U4 ∪ U5 ∪ U6 (順序あり: 軸→相手)
    #   馬連 = R1 ∪ R2 ∪ R3               (順序なし)
    #   各条件: 軸(_GZ_AXIS_IDX) → 紐(_GZ_PARTNER_FN)。
    #   重複買い目の条件タグは「賭け金が最大の条件」を採用する(先着優先)。
    #   ★旧v135の参考馬連カスケード(参考①②/F00補充)は全廃。
    #     馬連は採用条件(R1-R4)で直接発火させるため、連系0点レースは
    #     「買い目なし」として素直に見送る。
    # ══════════════════════════════════════════════════════════════
    gz_umatan_cond = {}   # (ban1,ban2) -> 条件名
    gz_umaren_cond = {}   # (ban1,ban2) -> 条件名
    gz_wide_cond = {}     # ★v136_001 (ban1,ban2) -> 条件名(ワイド)

    # ★v135_004: 重複買い目のタグ優先順は Kelly比率ではなくレジストリの order
    #   (選抜順位: 検証的中数と学習/検証の的中率一致で序列化)。均等買いでは
    #   金額差が無いため、タグは「より信頼できる条件」を先着させる。
    _GZ_UMATAN = sorted(GZ_UMATAN_NAMES, key=lambda c: GZ_COND_DEFS[c]['order'])
    _GZ_UMAREN = sorted(GZ_UMAREN_NAMES, key=lambda c: GZ_COND_DEFS[c]['order'])
    _GZ_WIDE = sorted(GZ_WIDE_NAMES, key=lambda c: GZ_COND_DEFS[c]['order'])  # ★v136_001

    # ★v135_010: 上流データの結合ミス等で 番 が重複すると、軸と同じ馬番の
    #   相手が混ざり『1→1』という買えない買い目が生成されていた。
    #   _partners_* は DataFrame index で軸を除いているが馬番の重複は見ていないため、
    #   買い目生成側で馬番一致を弾き、重複を検出したら一度だけ警告する。
    _dup_bans = [int(b) for b, c in group['番'].value_counts().items() if c > 1]
    if _dup_bans:
        print(f'  [警告] 馬番が重複しています: {sorted(_dup_bans)} '
              f'(自己ペアの買い目は除外しますが、出馬表データを確認してください)')

    def _gz_collect(names, pair_index, ordered):
        """条件名リストから買い目を収集。
        pair_index: 0=馬連 / 1=馬単 / 2=ワイド   ordered: 馬単は順序あり(True)"""
        _recs = []; _seen = set(); _tag = {}
        for _cname in names:
            _cax = _GZ_AXIS_IDX.get(GZ_COND_DEFS[_cname]['axis'])
            if _cax is None:
                continue
            _pidx = _GZ_PARTNER_FN[_cname](_cax)
            if not _pidx:
                continue
            _p_bans = {int(group.loc[_i, '番']) for _i in _pidx}
            _all = _mk_pairs(_cax, _wp_order)[pair_index]
            for r in _all:
                if r.ban2 not in _p_bans:
                    continue
                if r.ban1 == r.ban2:      # ★v135_010: 馬番重複による自己ペアを除外
                    continue
                _key = (r.ban1, r.ban2)
                _dk = _key if ordered else frozenset(_key)
                if _dk in _seen:
                    continue
                _seen.add(_dk)
                _recs.append(r)
                _tag[_key] = _cname
        return _recs, _tag

    umatan_recs, gz_umatan_cond = _gz_collect(_GZ_UMATAN, 1, True)
    umaren_recs, gz_umaren_cond = _gz_collect(_GZ_UMAREN, 0, False)
    # ★v136_001: ワイドも本線枠として同じ経路で生成する(_mk_pairs の3番目)。
    wide_recs, gz_wide_cond = _gz_collect(_GZ_WIDE, 2, False)

    # ══════════════════════════════════════════════════════
    # ★v136_004: 参考予想 — 軸=◎(アンサンブル単勝確率1位)。
    #   ・相手 = 『複勝確率47.7%以上 または 紐馬指数60以上』の該当馬“全員”
    #     (_himo_all_idxs。該当0頭のときのみ紐馬指数上位3頭)。
    #     ★予想印▲△☆は記号が3種類しかないため上位3頭までの表示だが、
    #       参考予想の相手は印の数に関係なく該当馬全員を対象にする
    #       (紐馬指数60以上の馬が印なしで買い目から漏れないようにする)。
    #   ・券種は馬連・ワイド(REF_HIMO_BET_TYPES)。
    #   ・実弾には一切関与しない(gz_stakes / 財布を通らない)。
    #   ・本線枠が既に買っている買い目は除外する。
    # ══════════════════════════════════════════════════════
    def _build_reference_himo(_excl):
        """◎(軸) × ○(アンサンブル単勝確率2位) および
        ◎ × 『複勝確率47.7%以上 or 紐馬指数60以上』の該当馬全員(_himo_all_idxs)
        への参考買い目(馬連・ワイド)。◎×○を先頭に、続いて紐馬候補を並べる。

        戻り値は [{kind, name, axis_ban, partner_bans, stats, is_fallback, desc}, ...]。
        券種(kind)ごとに1つのdictへ相手馬番をまとめる。
        """
        _ax = axis_idx
        if _ax is None:
            return []
        try:
            _ax_ban = int(group.loc[_ax, '番'])
        except Exception:
            return []

        # ○(second_idx)の馬番を取得
        _ni_ban = None
        if second_idx is not None and second_idx != _ax:
            try:
                _ni_ban = int(group.loc[second_idx, '番'])
            except Exception:
                pass

        if not _himo_all_idxs and _ni_ban is None:
            return []

        if _ref_new_partners is not None:
            _desc = '◎→○▲△(配当の妙味を重視した上位3頭)'
        else:
            _desc = ('◎○+複勝確率%.1f%%以上 or 紐馬指数%.0f以上' % (HIMO_PLACE_PROB_MIN, HIMO_HIMOBA_IDX_MIN)
                     if not _himo_is_fallback
                     else '◎○+該当馬なし→紐馬指数上位3頭(フォールバック)')

        _out = []
        for _kd in REF_HIMO_BET_TYPES:
            _p_bans = []
            # ◎×○を先頭に追加
            if _ni_ban is not None:
                _key_ni = (_kd,) + tuple(sorted((_ax_ban, _ni_ban)))
                if _key_ni not in (_excl or set()):
                    _p_bans.append(_ni_ban)
            # ◎×紐馬候補(○を除く)を追加
            for _i in _himo_all_idxs:
                try:
                    _p_ban = int(group.loc[_i, '番'])
                except Exception:
                    continue
                if _p_ban == _ax_ban:
                    continue
                if _ni_ban is not None and _p_ban == _ni_ban:
                    continue                          # ○は既に先頭に追加済み
                _key = (_kd,) + tuple(sorted((_ax_ban, _p_ban)))
                if _key in (_excl or set()):         # 本線枠が既に買っている
                    continue
                _p_bans.append(_p_ban)
            if not _p_bans:
                continue
            _out.append(dict(kind=_kd, name='', axis_ban=_ax_ban,
                             partner_bans=_p_bans, stats={},
                             is_fallback=_himo_is_fallback, desc=_desc))
        return _out

    reference_bets = []
    try:
        _ref_excl = set()
        for _k in gz_umatan_cond:              # (ban1, ban2)
            _ref_excl.add(('umatan', _k[0], _k[1]))
        for _k in gz_umaren_cond:              # frozenset({b1, b2})
            _ref_excl.add(('umaren',) + tuple(sorted(_k)))
        for _k in gz_wide_cond:
            _ref_excl.add(('wide',) + tuple(sorted(_k)))
        # ★v136_010: 新ロジックでは既定で本線枠との重複を除かない(常に◎-○▲△の3点)
        if _ref_new_partners is not None and not REF_EXCLUDE_MAIN:
            _ref_excl = set()
        reference_bets = _build_reference_himo(_ref_excl)
    except Exception as _e:
        # 参考予想でレースを落とさない。空にして続行する。
        print(f'[参考予想] 生成に失敗したためスキップ: '
              f'{type(_e).__name__}: {_e}')
        reference_bets = []

    # ★v136_003: 旧v135_014の『発火した買い目のヒモ優先で○▲△☆を並べ替える』
    #   ロジックは廃止。印(◎○▲△☆)は上で確定済み(◎○=アンサンブル単勝確率
    #   1位/2位、▲△☆=複勝確率47.7%以上or紐馬指数60以上、または紐馬指数上位3頭)
    #   のままとし、買い目の発火有無によって印を並べ替えることはしない。

    # ★v136_001: wide_recs は上の _gz_collect(_GZ_WIDE, 2, False) で生成済み。
    #   (旧v135では『ワイドは検証で優位なしのため出さない』として空固定していたが、
    #    最適紐008 の再分析で軸40-49.9%帯のワイドが的中54.0%/回収104.1%と
    #    本線基準を満たしたため、本線枠の券種として復活させた。ここで
    #    上書きすると生成結果が消えるので絶対に空代入を戻さないこと。)
    trio_recs = []

    # ══════════════════════════════════════════════════════════════
    # ★v135_014b(要望): 買い目なしレースの補充馬連(参考)。
    #   条件: 馬単・馬連ともに発火0点。
    #   内容: ◎(単勝確率1位)から、○▲△☆の4頭のうち『単勝期待値』上位3頭
    #         への馬連を表示する。3頭未満ならある分だけ。
    #   扱い: 統計的裏付けなし → タグ '参考'。yuuryoku/穴 には計上しない。
    #         後段の force_reference と同様、備考で参考である旨が出る。
    # ══════════════════════════════════════════════════════════════
    _ref_umaren_pairs = set()          # (ban1,ban2) 順序なし判定用
    _has_real_conn = bool(umatan_recs or umaren_recs or wide_recs)   # 補充前の実発火有無
    # ★v135_022: 補充馬連を既定オフにした。REF_UMAREN_ENABLE で戻せる。
    #   2026/08/10 の実出力78点で 期待回収率(オッズ×的中率) 平均77.0%、
    #   100%超が 0/78点。表示している数字自体が全点マイナスを示していた。
    #   ランク族探索でも馬連は全条件で検証マージン -5.0〜-7.4pt であり、
    #   補充が走る『軸wp<40』帯はまさに馬連の検証ROIが71〜75%の領域。
    #   役割は参考枠のフォールバック(馬単FB)と重複するため、
    #   唯一プラスの証拠がある券種=馬単 側に一本化する。
    if REF_UMAREN_ENABLE and not umatan_recs and not umaren_recs and not wide_recs:
        # ══════════════════════════════════════════════════════════
        # ★v135_015: 補充馬連3点ロジック(全面刷新)
        #
        #   旧版は「○▲△☆のうち単勝期待値上位3頭」という統計的裏付けの
        #   ない選び方だった。本版は prob_model_conditions_with_axis.json の
        #   検証結果から相手の優先順位を決める。
        #
        #   ★重要: 補充が走るのは『本線枠もカバー枠も発火しなかった』
        #     レース、すなわち軸の単勝確率が概ね40%未満の帯である。
        #     したがって軸wp>=40/50 帯の検証結果は適用できない。
        #     以下は軸wp>=20〜30帯(=まさにこの帯)の馬連検証から採った。
        #
        #   優先順位(いずれも軸=単勝確率1位=◎ からの馬連):
        #     P1 複勝確率_生>=30 かつ EV_生 1.2-2
        #        → 検証N=919 / ROI101.5% / 的中12.7% (軸wp>=30)
        #          この帯で最もNが厚く、唯一100%を超える実用的な組み合わせ。
        #     P2 単指数>=70 かつ EV_生 1.5+
        #        → 検証N=515 / ROI102.0% / 的中11.5% (軸wp>=30)
        #     P3 単指数>=70 かつ 騎手指数>=40
        #        → 検証N=1076 / ROI98.4% / 的中17.0% (軸wp>=30)
        #          ROIは100%割れだがこの帯で最も的中率が高い。穴埋め役。
        #     P4 上記で3点に満たない分を 複勝確率_生 降順で補充。
        #
        #   ★扱いは従来どおり『参考』。上記ROIはいずれもCI下限未算定で、
        #     100%前後という水準は控除率を考えると勝てる根拠にならない。
        #     買い目を空にせず情報を出すための表示であり、投資ではない。
        # ══════════════════════════════════════════════════════════
        _ref_cands = [_i for _i in _wp_order if _i != axis_idx]

        def _p1(_i):
            return _ge(_i, '複勝確率_生', 30.0) and _ev_in(_i, 1.2, 2.0)

        def _p2(_i):
            return _ge(_i, '単指数', 70.0) and _ev_in(_i, 1.5)

        def _p3(_i):
            return _ge(_i, '単指数', 70.0) and _ge(_i, '騎手指数', 40.0)

        def _fuku_of(_i):
            _v = _h_val(_i, '複勝確率_生')
            return _v if _v is not None else -1.0

        _ref_partners = []
        _ref_tier_of = {}
        for _tag, _fn in (('P1', _p1), ('P2', _p2), ('P3', _p3)):
            if len(_ref_partners) >= 3:
                break
            # 同一優先度内は複勝確率_生 降順(同値は馬番昇順)。
            _hit = sorted([_i for _i in _ref_cands
                           if _i not in _ref_partners and _fn(_i)],
                          key=lambda _i: (-_fuku_of(_i), int(group.loc[_i, '番'])))
            for _i in _hit:
                if len(_ref_partners) >= 3:
                    break
                _ref_partners.append(_i)
                _ref_tier_of[_i] = _tag
        if len(_ref_partners) < 3:
            _rest = sorted([_i for _i in _ref_cands if _i not in _ref_partners],
                           key=lambda _i: (-_fuku_of(_i), int(group.loc[_i, '番'])))
            for _i in _rest[:3 - len(_ref_partners)]:
                _ref_partners.append(_i)
                _ref_tier_of[_i] = 'P4'

        if _ref_partners:
            _ref_ur = _mk_pairs(axis_idx, _ref_partners)[0]
            for r in _ref_ur:
                if r.ban1 == r.ban2:
                    continue
                _key = (r.ban1, r.ban2)
                if frozenset(_key) in {frozenset(k) for k in _ref_umaren_pairs}:
                    continue
                umaren_recs.append(r)
                # どの優先度で拾ったかをタグに残す(集計で効果を測るため)。
                _tg = 'P4'
                for _i in _ref_partners:
                    if int(group.loc[_i, '番']) == r.ban2:
                        _tg = _ref_tier_of.get(_i, 'P4')
                        break
                gz_umaren_cond[_key] = f'参考{_tg}'
                _ref_umaren_pairs.add(_key)

    # 連系買い目の有無
    _has_conn = bool(umatan_recs or umaren_recs or wide_recs)


    # 有力/妙味フラグ集合(表示・備考用)
    # ★変更(要望): 「有力」は検証済み馬単条件
    #   [axis_win_prob>=50 & axis_kishu_idx>=40] + tan_idx>=50 & kishu_idx>=50 のみに付与する。
    #   馬連の有力判定は廃止(常に参考)。単勝も同様(上記)。
    #   ※検証成績(2026/03-06の4ヶ月・完全アウトオブサンプル):
    #       軸kishu>=40あり  検証ROI 100.9% (R=685, 的中163, CI下限70.7%)
    #       軸kishu条件なし  検証ROI  95.8% (R=808, 的中191, CI下限68.5%)
    #     月別は 3月81.7% / 4月77.8% / 5月76.3% / 6月178.3% と6月のみ突出。
    #     ブートストラップCI下限は100%を割っており、統計的に利益が確認された
    #     条件ではない。あくまで「最も的中密度が高く、少数の大穴に依存しない」
    #     候補としての表示である点に注意。
    yuuryoku_umaren = set()
    # ★v135_011: 「有力」は本線枠(tier='main')の馬単のみに付与する。
    #   紐カバー枠は低的中率/下限不足のため『参考』扱いのままにする。
    yuuryoku_umatan = set((r.ban1, r.ban2) for r in umatan_recs
                          if gz_tier(gz_umatan_cond.get((r.ban1, r.ban2), '')) == 'main')
    yuuryoku_wide = set()   # ★v136_001: 本線ワイド確定後に下段で再設定する
    myomi_umaren = set()
    myomi_umatan = set()
    anaba_umaren = []; anaba_umatan = []; anaba_wide = []
    is_invest = False

    # ── 開催地/信頼度ゲート(参考強制) ──
    try:
        _venue = str(group['場所'].iloc[0]).strip() if '場所' in group.columns else ''
    except Exception:
        _venue = ''
    _venue_skip = _venue in NSL_SKIP_VENUES
    if _venue_skip:
        umaren_recs = []; umatan_recs = []; wide_recs = []
        yuuryoku_umaren = set(); yuuryoku_umatan = set()
        # ★買い目をクリアしたら条件タグもクリア(賭け金計算・発火表示が
        #   クリア済み買い目に基づいて残る不整合を防ぐ)
        gz_umatan_cond = {}; gz_umaren_cond = {}
        gz_wide_cond = {}   # ★v136_001
        if tan_priority_bans:
            # 単勝の有力条件はゲートに優先 → 単勝は残し有力のまま
            pass
        else:
            tan_recs = []; yuuryoku_tan = set()

    # fav_bans(人気1-3の馬番)
    try:
        # ★人気0(欠損)を「人気1-3」に含めない
        _pop_s = pd.to_numeric(group['単勝人気'], errors='coerce') if '単勝人気' in group.columns else None
        fav_bans = (set(int(x) for x in group.loc[(_pop_s >= 1) & (_pop_s <= 3), '番'].dropna().astype(int).tolist())
                    if _pop_s is not None else set())
    except Exception:
        try:
            _ap = pd.to_numeric(group['adjusted_popularity'], errors='coerce')
            fav_bans = set(int(x) for x in group.loc[(_ap >= 1) & (_ap <= 3), '番'].dropna().astype(int).tolist())
        except Exception:
            fav_bans = set()

    # ── 期待水準(◎の儲かりやすさ=期待値ベースの解説) ──
    #   ★v133_modified: 「期待水準」は勝率の羅列ではなく、軸馬の期待値(EV)から
    #     単勝の妙味(市場オッズに対する優位性)を言葉で解説する。
    a_winP_disp = float(group.loc[axis_idx, '単勝確率'])
    a_fukuP_disp = float(group.loc[axis_idx, '複勝確率'])
    a_ev_disp = float(group.loc[axis_idx, '単勝期待値'])   # 単勝オッズ×勝率(基準100)
    a_odds_disp = (float(group.loc[axis_idx, '単勝オッズ'])
                   if '単勝オッズ' in group.columns
                   else float(group.loc[axis_idx, 'adjusted_odds']))
    # 単勝EV(基準100=トントン)を回収率イメージで区分
    if a_ev_disp >= 130:
        _ev_word = '非常に高い(妙味大)'
    elif a_ev_disp >= 110:
        _ev_word = '高い(妙味あり)'
    elif a_ev_disp >= 95:
        _ev_word = '標準的(ほぼ適正オッズ)'
    else:
        _ev_word = '低い(過剰人気ぎみ)'
    expected_win_rate = '単勝期待値%.0f(%s)' % (a_ev_disp, _ev_word)
    expected_place_rate = '勝率%.0f%%/複勝率%.0f%%・単勝オッズ%.1f倍' % (
        a_winP_disp, a_fukuP_disp, a_odds_disp)

    # ── コメント(軸馬の騎手・馬の総合解説) ──
    #   ★v133_modified: 「コメント」は軸馬そのものの総合解説(騎手/脚質/指数/人気)。
    try:
        a_name = str(group.loc[axis_idx, '馬名']) if '馬名' in group.columns else ''
    except Exception:
        a_name = ''
    try:
        a_kishu = str(group.loc[axis_idx, '騎手']) if '騎手' in group.columns else ''
    except Exception:
        a_kishu = ''
    try:
        a_kishu_idx = float(group.loc[axis_idx, '騎手指数'])
    except Exception:
        a_kishu_idx = 0.0
    try:
        a_style = str(group.loc[axis_idx, '展開']) if '展開' in group.columns else ''
    except Exception:
        a_style = ''
    try:
        a_pop = (int(group.loc[axis_idx, '単勝人気']) if '単勝人気' in group.columns
                 else int(group.loc[axis_idx, 'adjusted_popularity']))
    except Exception:
        a_pop = 0
    a_fukuI = float(group.loc[axis_idx, '複指数'])
    # 騎手評価
    if a_kishu_idx >= 60:
        _k_word = '好騎手'
    elif a_kishu_idx >= 50:
        _k_word = '標準的な騎手'
    else:
        _k_word = 'やや割引く騎手'
    # 能力評価(単指数)
    if _axis_tan_idx >= 90:
        _idx_word = '能力上位(単指数最上位クラス)'
    elif _axis_tan_idx >= 70:
        _idx_word = '能力上位'
    else:
        _idx_word = '能力は平均域'
    _cmt_parts = []
    if a_name:
        _cmt_parts.append('◎%d番 %s' % (axis_ban_wp, a_name))
    else:
        _cmt_parts.append('◎%d番' % axis_ban_wp)
    if a_kishu:
        _cmt_parts.append('鞍上%s(%s)' % (a_kishu, _k_word))
    _cmt_parts.append('%s。単指数%.0f/複指数%.0f' % (_idx_word, _axis_tan_idx, a_fukuI))
    if a_style:
        _cmt_parts.append('脚質%s' % a_style)
    if a_pop:
        _cmt_parts.append('%d番人気' % a_pop)
    _cmt_parts.append('単勝確率%.1f%%・複勝確率%.1f%%' % (a_winP_disp, a_fukuP_disp))
    judgment_comment = ' / '.join(_cmt_parts)
    if _venue_skip:
        judgment_comment = '【見送り開催地:%s】 ' % _venue + judgment_comment

    # 信頼度区分(既存formula流用) → 参考強制判定
    _conf_tier_label = ''
    _tier_ref = False
    try:
        _probe = {'axis_win_prob': a_winP_disp, 'axis_place_prob': a_fukuP_disp,
                  'axis_zscore': 0.0, 'judgment_class': 'C軸'}
        _conf_tier_label = calculate_confidence_score(_probe).get('category', '')
        _tier_ref = _conf_tier_label in NSL_TIER_REFERENCE_ONLY
    except Exception:
        _conf_tier_label = ''; _tier_ref = False
    if _tier_ref and not _venue_skip:
        judgment_comment += ' 【信頼度%s区分: 参考】' % _conf_tier_label

    # 参考強制(フォールバック時 or 単勝が参考 or 信頼度/開催地ゲート)
    #   ★v135_014b: 補充馬連(参考)しか無いレースも『参考』強制にする。
    #     _has_conn は補充後 True になるが、中身が参考のみなら賭け金は0円表示。
    _only_ref_conn = bool(_ref_umaren_pairs) and not _has_real_conn
    force_reference = ((_venue in NSL_REFERENCE_ONLY_VENUES) or _tier_ref
                       or (not _has_conn) or _only_ref_conn)

    # ══════════════════════════════════════════════════════════
    # ★v134: 安定運用(トリガミ覚悟)の単勝/複勝 買い目判定
    #   軸馬(単勝確率1位)が4条件を満たすとき、開催地ゲートに応じて
    #   単勝/複勝を『安定』(ゲート内) or 『参考(安定外)』(ゲート外)で出す。
    # ══════════════════════════════════════════════════════════
    stable_tan_rec = None    # SingleRec or None
    stable_fuku_rec = None
    stable_tan_tag = ''
    stable_fuku_tag = ''
    if (STABLE_BET_ENABLE and _stable_axis_ban is not None
            and _stable_axis_ok(_stable_win_prob, _stable_place_prob, _stable_raw_pop)):
        # 表示用の的中率は生確率(検証と同一基準)。オッズ/EVは軸行から。
        try:
            _ax_odds = float(group.loc[_tan_target_idx, 'adjusted_odds'])
        except Exception:
            _ax_odds = 0.0
        # 単勝安定レコード(的中率=単勝確率_生)
        stable_tan_rec = SingleRec(
            ban=_stable_axis_ban,
            hit_rate=round(float(_stable_win_prob), 1),
            odds=_ax_odds,
            ev=int(round(float(group.loc[_tan_target_idx, '単勝期待値']))
                   if '単勝期待値' in group.columns else 0))
        # 複勝安定レコード(的中率=複勝確率_生)。複勝オッズは概算(なければ0)。
        try:
            _ax_fodds = round(1.0 + (100.0 / max(float(_stable_place_prob), 1.0) - 1.0)
                              * 0.30, 1)   # 複勝オッズの粗い目安(下限1.0)
            _ax_fodds = max(1.0, _ax_fodds)
        except Exception:
            _ax_fodds = 0.0
        stable_fuku_rec = SingleRec(
            ban=_stable_axis_ban,
            hit_rate=round(float(_stable_place_prob), 1),
            odds=_ax_fodds, ev=0)
        # 開催地ゲート
        stable_tan_tag = '安定(単)' if _venue in STABLE_TAN_VENUES else '参考(安定外)'
        stable_fuku_tag = '安定(複)' if _venue in STABLE_FUKU_VENUES else '参考(安定外)'

    # ══════════════════════════════════════════════════════════
    # ★v135_025: 次点本線枠 (複勝【安定】★採用 軸 × 複勝確率_生50%以上の最上位1頭)
    #   馬連/ワイドを各1点。8/1-8/23 実測(59R中39R発火):
    #     馬連 39点 的中46.2% ROI125.9% / ワイド 39点 的中76.9% ROI106.7%
    #   ※N=39。CI下限は100%を大きく割るため『検証中の候補』の位置づけ。
    #   発火条件:
    #     ① 複勝【安定】が ★採用(ゲート内 = 安定(複)) であること
    #     ② 軸を除く出走馬に 複勝確率_生 >= SUBLINE_PARTNER_MIN_PLACE の馬がいること
    #     ③ 相手 = ②のうち 複勝確率_生 が最大の1頭(同値は元の並び順)
    # ══════════════════════════════════════════════════════════
    subline_rec = None
    if (SUBLINE_LEGACY_ENABLE
            and stable_fuku_rec is not None
            and stable_fuku_tag.startswith('安定')
            and _stable_axis_ban is not None):
        _sl_best_idx = None
        _sl_best_pp = None
        for _i in group.index:
            if _i == _tan_target_idx:
                continue
            _pp = _h_val_g(_i, '複勝確率_生')
            if _pp is None or _pp < SUBLINE_PARTNER_MIN_PLACE:
                continue
            if _sl_best_pp is None or _pp > _sl_best_pp:
                _sl_best_pp = _pp
                _sl_best_idx = _i
        if _sl_best_idx is not None:
            try:
                _sl_ban = int(group.loc[_sl_best_idx, '番'])
            except Exception:
                _sl_ban = None
            if _sl_ban is not None:
                try:
                    _sl_odds = float(group.loc[_sl_best_idx, 'adjusted_odds'])
                except Exception:
                    _sl_odds = 0.0
                subline_rec = {
                    'axis_ban': int(_stable_axis_ban),
                    'partner_ban': _sl_ban,
                    'axis_place': round(float(_stable_place_prob), 1),
                    'partner_place': round(float(_sl_best_pp), 1),
                    'partner_odds': _sl_odds,
                }

    # ══════════════════════════════════════════════════════════
    # ★v135_026: 次点本線枠B (軸wp40+ × 相手複指数24以上 × 10頭以下)
    #   馬連/ワイドを該当馬ぶん。8/1-8/23 実測(74点):
    #     馬単 13.5%/181.4% ・ 馬連 14.9%/161.8% ・ ワイド 31.1%/111.6%
    #   ※11頭以上では馬単73.0%と逆転するため頭数上限が必須。
    #     N=74・CI下限100%未満のため『検証中の候補』の位置づけ。
    #   発火条件:
    #     ① 出走頭数 <= SUBLINE_B_MAX_FIELD
    #     ② 軸 = 単勝確率_生 >= SUBLINE_B_AXIS_MIN_WP のうち最大の1頭
    #     ③ 相手 = 軸以外で 複指数 >= SUBLINE_B_PARTNER_MIN_FUKU の全馬
    #        (複指数降順 → 同値は複勝確率_生 降順)
    # ══════════════════════════════════════════════════════════
    subline_b_rec = None
    if SUBLINE_LEGACY_ENABLE and len(group) <= SUBLINE_B_MAX_FIELD:
        _sb_axis_idx = None
        _sb_axis_wp = None
        for _i in _wp_order:                      # 単勝確率_生 降順
            _wv = _h_val_g(_i, '単勝確率_生')
            if _wv is not None and _wv >= SUBLINE_B_AXIS_MIN_WP:
                _sb_axis_idx = _i
                _sb_axis_wp = _wv
                break
        if _sb_axis_idx is not None:
            _sb_partners = []
            for _i in group.index:
                if _i == _sb_axis_idx:
                    continue
                _fi = _h_val_g(_i, '複指数')
                if _fi is None or _fi < SUBLINE_B_PARTNER_MIN_FUKU:
                    continue
                try:
                    _pb = int(group.loc[_i, '番'])
                except Exception:
                    continue
                try:
                    _po = float(group.loc[_i, 'adjusted_odds'])
                except Exception:
                    _po = 0.0
                _sb_partners.append({
                    'ban': _pb, 'fuku_idx': round(float(_fi), 1),
                    'odds': _po,
                    'place': _h_val_g(_i, '複勝確率_生') or 0.0,
                })
            if _sb_partners:
                _sb_partners.sort(key=lambda _d: (-_d['fuku_idx'], -_d['place']))
                try:
                    _sb_axis_ban = int(group.loc[_sb_axis_idx, '番'])
                except Exception:
                    _sb_axis_ban = None
                if _sb_axis_ban is not None:
                    try:
                        _sb_axis_odds = float(group.loc[_sb_axis_idx, 'adjusted_odds'])
                    except Exception:
                        _sb_axis_odds = 0.0
                    subline_b_rec = {
                        'axis_ban': _sb_axis_ban,
                        'axis_wp': round(float(_sb_axis_wp), 1),
                        'axis_odds': _sb_axis_odds,
                        'field': int(len(group)),
                        'partners': _sb_partners,
                    }

    # ══════════════════════════════════════════════════════════
    # ★v135_027: 次点本線枠C (軸wp50+ × 相手 複勝確率40+ & 騎手指数50+)
    #   有望条件.json の最良骨格(J01/J16)。馬単(軸→相手)・馬連を該当馬ぶん。
    #   発火条件:
    #     ① 軸 = 単勝確率_生 >= SUBLINE_C_AXIS_MIN_WP のうち最大の1頭
    #     ② 相手 = 軸以外で 複勝確率_生 >= SUBLINE_C_PARTNER_MIN_PLACE
    #        かつ 騎手指数 >= SUBLINE_C_PARTNER_MIN_KISHU の全馬
    #        (単勝確率_生 降順 → 同値は複勝確率_生 降順)
    #   ※sign通過だが CI下限100%未満 → 『検証中の候補』。本線枠には入れない。
    # ══════════════════════════════════════════════════════════
    subline_c_rec = None
    _sc_axis_idx = None
    _sc_axis_wp = None
    for _i in (_wp_order if SUBLINE_LEGACY_ENABLE else []):   # 単勝確率_生 降順
        _wv = _h_val_g(_i, '単勝確率_生')
        if _wv is not None and _wv >= SUBLINE_C_AXIS_MIN_WP:
            _sc_axis_idx = _i
            _sc_axis_wp = _wv
            break
    if _sc_axis_idx is not None:
        _sc_partners = []
        for _i in _wp_order:                   # 単勝確率_生 降順で相手を集める
            if _i == _sc_axis_idx:
                continue
            _pp = _h_val_g(_i, '複勝確率_生')
            _kj = _h_val_g(_i, '騎手指数')
            if _pp is None or _pp < SUBLINE_C_PARTNER_MIN_PLACE:
                continue
            if _kj is None or _kj < SUBLINE_C_PARTNER_MIN_KISHU:
                continue
            try:
                _pb = int(group.loc[_i, '番'])
            except Exception:
                continue
            try:
                _po = float(group.loc[_i, 'adjusted_odds'])
            except Exception:
                _po = 0.0
            _sc_partners.append({
                'ban': _pb, 'place': round(float(_pp), 1),
                'kishu': round(float(_kj), 1), 'odds': _po,
                '_wp': _h_val_g(_i, '単勝確率_生') or 0.0,
            })
        if _sc_partners:
            _sc_partners.sort(key=lambda _d: (-_d['_wp'], -_d['place']))
            try:
                _sc_axis_ban = int(group.loc[_sc_axis_idx, '番'])
            except Exception:
                _sc_axis_ban = None
            if _sc_axis_ban is not None:
                try:
                    _sc_axis_odds = float(group.loc[_sc_axis_idx, 'adjusted_odds'])
                except Exception:
                    _sc_axis_odds = 0.0
                subline_c_rec = {
                    'axis_ban': _sc_axis_ban,
                    'axis_wp': round(float(_sc_axis_wp), 1),
                    'axis_odds': _sc_axis_odds,
                    'partners': _sc_partners,
                }

    # ══════════════════════════════════════════════════════════
    # ★v135_028: 次点本線枠(勝ち筋マップ)生成
    #   本線(MB1/MB2)に採用しなかった sign通過条件を表示専用で発火させる。
    #   ・同じ買い目は本線を優先して1つに集約(本線が持つ買い目・次点内の重複を除く)。
    #   ・実弾ではない(gz_stakes/財布を通らない)。CI下限100%未満の検証候補。
    #   ・馬単=順序あり / 馬連・ワイド=順序なし(frozenset)で重複排除。
    # ══════════════════════════════════════════════════════════
    # ★v135_030: 軸コードの汎用パーサ(AO軸・複勝率軸・騎手軸を追加)。
    #   'w{n}'      → 単勝確率_生>=n の最上位1頭 (例 w50 / w40 / w20)
    #   'pp{n}'     → 複勝確率_生>=n            (例 pp60)
    #   'ks{n}'     → 騎手指数>=n(補正前)        (例 ks50)
    #   'ao{t}_{o}' → 単指数>=t & 単勝オッズ<=o/10 の最上位1頭 (例 ao90_18=単90&オッズ1.8)
    #   ※軸の「最上位」は他条件と同一(単勝確率_生 最大)。単勝オッズはAI予測オッズ。
    def _jiten_axis(_axkind):
        if not _axkind:
            return None
        # ★v136_018: 実戦条件シートの観察条件の軸(本線と同じ選び方)
        if _axkind == 'S:ti90max_o19':
            return _gz_axis_s_ti90_o19
        if _axkind == 'S:aw50max_o19':
            return _gz_axis_s_aw50_o19
        try:
            if _axkind.startswith('ao'):
                # ★v135_031: 検証(hc8_fdr の _per_race_payoff_axis)と厳密一致させる。
                #   検証は「単指数>=t を満たす馬のうち単勝確率_生 最大の1頭」を軸に選び、
                #   その軸馬の補正前オッズが o を超えたら『そのレースは見送り』(下位馬に
                #   繰り上げない)。以前は tan&odds 両方を満たす最上位へ繰り上げていたが、
                #   それだと検証が見送るレースで別の軸馬を買ってしまい定義がズレる。
                _a, _b = _axkind[2:].split('_')
                _tan_thr = float(_a); _odds_max = float(_b) / 10.0
                _ax = _axis_by_col('単指数', _tan_thr)   # tan>=t のうち単勝確率_生 最大
                if _ax is None:
                    return None
                _o = _h_val_g(_ax, '単勝オッズ')          # 補正前(AI予測)単勝オッズ
                if _o is None or _o > _odds_max:         # 軸オッズ足切り: 超過は見送り
                    return None
                return _ax
            if _axkind.startswith('ks'):
                return _axis_by_col('騎手指数', float(_axkind[2:]))
            if _axkind.startswith('ti'):
                # ★v135_035: 単指数のみの軸(オッズ足切りなし)。おすすめ#29『単指数90以上』用。
                return _axis_by_col('単指数', float(_axkind[2:]))
            if _axkind.startswith('wa'):
                # ★v135_035: 補正後 単勝確率>=n の軸(おすすめランキング用/'w'より先に判定)。
                # ★v136_001: 本線枠(adj_win>=50)と同じ選び方に統一
                #   (補正後 単勝確率が最大の1頭。旧実装は _wp_order=生 降順の先頭ヒット)。
                return _axis_adj_win(float(_axkind[2:]))
            if _axkind.startswith('wb'):
                # ★v136_001: 補正後 単勝確率が [n, 50) の帯にある馬のうち最大の1頭。
                #   'wa50'(50%以上)とは排他の帯。最適紐008の軸『単勝確率40〜49.9%』。
                return _axis_adj_win_band(float(_axkind[2:]), 50.0)
            if _axkind.startswith('pp'):
                return _axis_by_col('複勝確率_生', float(_axkind[2:]))
            if _axkind.startswith('w'):
                return _axis_by_winprob(float(_axkind[1:]))
        except Exception:
            return None
        return None

    def _jiten_place_rank(_axis):
        """軸を除き 複勝確率_生 降順で 1..N の順位を返す(複勝上位N判定用)。
        並び・tie-break は _partners_S2 と一致(複勝確率_生 降順 → 馬番昇順)。"""
        def _pp(_i):
            _v = _h_val_g(_i, '複勝確率_生')
            return _v if _v is not None else -1.0
        _cand = sorted((_i for _i in group.index if _i != _axis),
                       key=lambda _i: (-_pp(_i), int(group.loc[_i, '番'])))
        return {_i: _r + 1 for _r, _i in enumerate(_cand)}

    def _jiten_atom_ok(_i, _atom, _place_rank=None):
        _col, _thr, _lab = _atom
        if _col == '_odds_le':             # ★v136_018: 表示の単オッズ(騎手補正後) <= _thr
            _o = _h_val_g(_i, 'adjusted_odds')
            return _o is not None and 0 < _o <= _thr
        if _col.startswith('S:'):          # ★v136_018: 表示値(小数1桁)で >= 判定
            return _sheet_ge(_i, _col[2:], _thr)
        if _col == '_pop':                 # 補正前 単勝人気 <= _thr
            return _pop_le(_i, _thr)
        if _col == '_place_rank':          # 複勝確率_生 上位 _thr 頭(軸除く)
            return _place_rank is not None and _place_rank.get(_i, 10 ** 9) <= _thr
        return _ge(_i, _col, _thr)

    def _jiten_ax_label(_axkind):
        if _axkind == 'S:ti90max_o19':      # ★v136_018
            return '軸[単指数90+の最大&予想オッズ1.9以下]'
        if _axkind == 'S:aw50max_o19':
            return '軸[単勝確率50%+の最大&予想オッズ1.9以下]'
        if _axkind and _axkind.startswith('ao'):
            try:
                _a, _b = _axkind[2:].split('_')
                return f'軸[単指数{int(float(_a))}+&オッズ{float(_b) / 10.0:g}以下]'
            except Exception:
                return _axkind
        if _axkind and _axkind.startswith('ti'):   # ★v135_035: 単指数のみ軸
            try:
                return f'軸[単指数{int(float(_axkind[2:]))}+]'
            except Exception:
                return _axkind
        if _axkind and _axkind.startswith('wa'):    # ★v135_035: 補正後 単勝確率 軸
            try:
                return f'軸[補正後勝率{float(_axkind[2:]):g}+]'
            except Exception:
                return _axkind
        if _axkind and _axkind.startswith('wb'):    # ★v136_001: 補正後 単勝確率の帯
            try:
                return f'軸[補正後勝率{float(_axkind[2:]):g}-49.9%]'
            except Exception:
                return _axkind
        return {'w40': '軸wp40+', 'w50': '軸wp50+', 'w20': '軸wp20+',
                'w25': '軸wp25+', 'w30': '軸wp30+',
                'pp50': '軸[複勝率50+]', 'pp60': '軸[複勝率60+]',
                'ks50': '軸[騎手50+]', 'ks60': '軸[騎手60+]'}.get(_axkind, _axkind)
    # 本線(MB1/MB2)が既に持つ買い目 → 次点から除外して集約する
    _main_ut_pairs = {(r.ban1, r.ban2) for r in umatan_recs}
    _main_ur_pairs = {frozenset((r.ban1, r.ban2)) for r in umaren_recs}
    _main_wd_pairs = {frozenset((r.ban1, r.ban2)) for r in wide_recs}   # ★v136_001
    _jiten_seen = set()        # (kind, dedup_key) 次点内の重複排除
    jiten_recs = []
    # ★v136_005で一度廃止し、★v136_018 で JITEN_ENABLE=True(実戦条件シートの観察2条件)に戻した。
    if JITEN_ENABLE:
        for _jd in JITEN_MAP_DEFS:
            _jax = _jiten_axis(_jd['axis'])
            if _jax is None:
                continue
            try:
                _jax_ban = int(group.loc[_jax, '番'])
            except Exception:
                continue
            _place_rank = _jiten_place_rank(_jax)   # ★v135_030: 複勝上位N判定用
            _jparts = []
            _jcands = [_i for _i in _wp_order              # 単勝確率_生 降順
                       if _i != _jax
                       and all(_jiten_atom_ok(_i, _a, _place_rank) for _a in _jd['atoms'])]
            if _jd.get('top1_himo'):       # ★v136_018: 該当馬のうち紐馬指数1位の1頭だけ
                _jcands = _top1_himo(_jcands)
            for _i in _jcands:
                try:
                    _pb = int(group.loc[_i, '番'])
                except Exception:
                    continue
                if _pb == _jax_ban:
                    continue
                _kind = _jd['kind']
                if _kind == 'umatan':
                    _dk = (_jax_ban, _pb)
                    if _dk in _main_ut_pairs:      # 本線馬単に集約済み
                        continue
                else:
                    _dk = frozenset((_jax_ban, _pb))
                    if _kind == 'umaren' and _dk in _main_ur_pairs:  # 本線馬連に集約済み
                        continue
                    if _kind == 'wide' and _dk in _main_wd_pairs:    # ★v136_001 本線ワイドに集約済み
                        continue
                _sk = (_kind, _dk)
                if _sk in _jiten_seen:             # 次点内の同一買い目を集約
                    continue
                _jiten_seen.add(_sk)
                try:
                    _po = float(group.loc[_i, 'adjusted_odds'])
                except Exception:
                    _po = 0.0
                _jparts.append({'ban': _pb, 'odds': _po,
                                'place': round(float(_h_val_g(_i, '複勝確率_生') or 0.0), 1)})
            if _jparts:
                jiten_recs.append({
                    'tag': _jd['tag'], 'kind': _jd['kind'],
                    'atoms_label': ' & '.join(_a[2] for _a in _jd['atoms']),
                    'axis_label': _jiten_ax_label(_jd['axis']),
                    'roi': _jd['roi'], 'hit': _jd['hit'], 'n': _jd['n'],
                    'src': _jd.get('src'), 'rank': _jd.get('rank'),
                    'typ': _jd.get('typ'),
                    'axis_ban': _jax_ban, 'partners': _jparts,
                })

    # ══════════════════════════════════════════════════════════
    # ★v135_033: 参考枠から『次点本線枠と重複する買い目』を除外する。
    #   本線枠との重複は build_reference_bets(exclude_pairs) で既に除外済みだが、
    #   次点本線枠(jiten_recs)は参考枠の生成より後に確定するため、ここで後追い除外する。
    #   ・馬単は順序あり(軸→相手)、馬連/ワイドは順序なし(frozenset)で突き合わせる。
    #   ・券種が異なる同一2頭(馬単a→b と 馬連a-b 等)は別の買い目なので除外しない。
    #   ・参考買い目の全点が次点と重複したら、その参考買い目自体を非表示にする。
    # ══════════════════════════════════════════════════════════
    # ★v136_010: 新ロジック(REF_EXCLUDE_MAIN=False)では次点本線枠との重複も除かない
    if reference_bets and jiten_recs and (_ref_new_partners is None or REF_EXCLUDE_MAIN):
        _jiten_pairs = set()
        for _jr in jiten_recs:
            _jkd = _jr['kind']; _jab = _jr['axis_ban']
            for _p in _jr['partners']:
                _pb = _p['ban']
                if _jkd == 'umatan':
                    _jiten_pairs.add(('umatan', _jab, _pb))
                else:                       # umaren / wide は順序なし
                    _jiten_pairs.add((_jkd, frozenset((_jab, _pb))))
        _kept_refs = []
        for _rf in reference_bets:
            _rkd = _rf.get('kind'); _rab = _rf.get('axis_ban')
            _kept_bans = []
            for _pb in (_rf.get('partner_bans') or []):
                if _rkd == 'umatan':
                    _dup = ('umatan', _rab, _pb) in _jiten_pairs
                else:
                    _dup = (_rkd, frozenset((_rab, _pb))) in _jiten_pairs
                if not _dup:
                    _kept_bans.append(_pb)
            if _kept_bans:                  # 残った点だけを表示。全点重複なら丸ごと非表示。
                _rf['partner_bans'] = _kept_bans
                _kept_refs.append(_rf)
        reference_bets = _kept_refs


    # ══════════════════════════════════════════════════════════
    # ★v135_004 賭け金 = 均等買い(1点 GZ_FLAT_UNIT 円) + 1レース点数上限
    #   ・Kelly配分/確信度重み/シュリンク/cap6% は全撤去。
    #     検証ROIは1点100円の均等買いで算出された値であり、それに合わせる。
    #   ・1レースの投下上限は 資金 × GZ_RACE_CAP_RATIO(既定0.5%)。
    #     上限点数 = 上限額 ÷ 1点単価(gz_race_max_points())。
    #   ・上限超過時は「条件のorder(選抜順位)が若い順 → 各条件内は
    #     相手の単勝確率_生 降順(=JSONの相手採用順)」で採用し、残りは切り捨てる。
    #     切り捨てた点数は gz_dropped_n に記録し備考へ出せるようにする。
    # ══════════════════════════════════════════════════════════
    _GZ_UNIT = GZ_FLAT_UNIT
    _gz_max_pts = gz_race_max_points()

    # ★v136_001: 券種→条件タグ辞書の対応(ワイド追加に伴い辞書引きへ)。
    def _cond_map_of(_kind):
        return {'umatan': gz_umatan_cond, 'umaren': gz_umaren_cond,
                'wide': gz_wide_cond}.get(_kind, {})

    def _gz_prio(_kind, _key, _pos):
        """(条件order, 元の並び順) — 小さいほど優先。"""
        _c = _cond_map_of(_kind).get(_key)
        return (GZ_COND_DEFS.get(_c, {}).get('order', 99), _pos)

    # ★v135_011: 本線枠(main)と紐カバー枠(ana)は別勘定で点数上限を掛ける。
    #   同一の上限を共有すると、発火点数の多い穴条件(A20/A25は相手が広い)が
    #   上限を食い潰し、的中率の高い本線条件が「見送り(点数上限)」になる。
    _gz_ana_max_pts = gz_ana_max_points()

    def _tier_of(_kind, _key):
        return gz_tier(_cond_map_of(_kind).get(_key, ''))

    # 全買い目を優先度順に並べ、枠ごとの上限点数で切る
    _gz_all = ([('umatan', (r.ban1, r.ban2), _p, r) for _p, r in enumerate(umatan_recs)]
               + [('umaren', (r.ban1, r.ban2), _p, r) for _p, r in enumerate(umaren_recs)]
               # ★v136_001: ワイドも同じ財布・同じ点数上限で扱う。
               + [('wide', (r.ban1, r.ban2), _p, r) for _p, r in enumerate(wide_recs)])
    _gz_all.sort(key=lambda t: _gz_prio(t[0], t[1], t[2]))
    _gz_main = [t for t in _gz_all if _tier_of(t[0], t[1]) == 'main']
    # ★v135_014b: 補充馬連(タグ'参考')は点数上限の対象外(発火0点の穴埋めのため
    #   賭け金は0円表示だが、表示は常に残す)。ana枠の点数も消費しない。
    # ★v135_015: 補充馬連のタグは '参考P1'〜'参考P4' になったため前方一致で判定。
    #   完全一致('参考')のままだと補充馬連が _gz_ana 側に流れ込み、
    #   カバー枠の点数上限を食い潰した上で 0円表示が残る不整合になる。
    def _is_ref(_kind, _key):
        _t = _cond_map_of(_kind).get(_key)
        return bool(_t) and str(_t).startswith('参考')

    _gz_ref = [t for t in _gz_all if _is_ref(t[0], t[1])]
    _gz_ana = [t for t in _gz_all if _tier_of(t[0], t[1]) != 'main'
               and not _is_ref(t[0], t[1])]
    _gz_keep = _gz_main[:_gz_max_pts] + _gz_ana[:_gz_ana_max_pts] + _gz_ref
    gz_dropped_n = max(0, len(_gz_main) - _gz_max_pts) + max(0, len(_gz_ana) - _gz_ana_max_pts)
    _keep_ut = {t[1] for t in _gz_keep if t[0] == 'umatan'}
    _keep_ur = {t[1] for t in _gz_keep if t[0] == 'umaren'}
    _keep_wd = {t[1] for t in _gz_keep if t[0] == 'wide'}   # ★v136_001

    if gz_dropped_n > 0:
        # 採用外の買い目は recs / 条件タグの両方から除去(0円表示を残さない)
        umatan_recs = [r for r in umatan_recs if (r.ban1, r.ban2) in _keep_ut]
        umaren_recs = [r for r in umaren_recs if (r.ban1, r.ban2) in _keep_ur]
        gz_umatan_cond = {k: v for k, v in gz_umatan_cond.items() if k in _keep_ut}
        gz_umaren_cond = {k: v for k, v in gz_umaren_cond.items() if k in _keep_ur}
        wide_recs = [r for r in wide_recs if (r.ban1, r.ban2) in _keep_wd]     # ★v136_001
        gz_wide_cond = {k: v for k, v in gz_wide_cond.items() if k in _keep_wd}
        yuuryoku_umatan = set((r.ban1, r.ban2) for r in umatan_recs
                              if gz_tier(gz_umatan_cond.get((r.ban1, r.ban2), '')) == 'main')

    # ★v136_001: ワイドの「有力」も本線枠(tier='main')に付与する。
    yuuryoku_wide = set((r.ban1, r.ban2) for r in wide_recs
                        if gz_tier(gz_wide_cond.get((r.ban1, r.ban2), '')) == 'main')

    _gz_active = sorted(set(gz_umatan_cond.values()) | set(gz_umaren_cond.values())
                        | set(gz_wide_cond.values()),
                        key=lambda c: GZ_COND_DEFS.get(c, {}).get('order', 99))
    # 全買い目を同額(均等買い)。
    # ★v135_010: 参考強制(開催地/信頼度ゲート)のレースは HTML/Excel/CSV 上
    #   『参考(見送り)』=0円と表示しながら gz_stakes には単価が入ったままで、
    #   gz_race_invest が実際には買わない分まで計上していた。日次の実投下額を
    #   集計する用途で過大計上になるため、賭け金側も0にする。
    _gz_bet_amt = 0 if force_reference else _GZ_UNIT
    gz_stakes = {}
    for _key in gz_umatan_cond:
        gz_stakes[(_key[0], _key[1], 'umatan')] = _gz_bet_amt
    for _key in gz_umaren_cond:
        gz_stakes[(_key[0], _key[1], 'umaren')] = _gz_bet_amt
    for _key in gz_wide_cond:                                  # ★v136_001
        gz_stakes[(_key[0], _key[1], 'wide')] = _gz_bet_amt
    gz_race_invest = _gz_bet_amt * len(gz_stakes)   # このレースの総投下額(円)

    # ★v135_011: 枠別の内訳(発火条件・点数・投下額)
    _gz_active_main = [c for c in _gz_active if gz_tier(c) == 'main']
    _gz_active_ana = [c for c in _gz_active if gz_tier(c) != 'main']
    _n_main = sum(1 for _k in gz_umatan_cond if gz_tier(gz_umatan_cond[_k]) == 'main') \
        + sum(1 for _k in gz_umaren_cond if gz_tier(gz_umaren_cond[_k]) == 'main') \
        + sum(1 for _k in gz_wide_cond if gz_tier(gz_wide_cond[_k]) == 'main')
    _n_ana = len(gz_stakes) - _n_main
    gz_main_invest = _gz_bet_amt * _n_main
    gz_ana_invest = _gz_bet_amt * _n_ana

    # marked_df は table_df から(印付き行)。fav_bans等は新ブロックで定義済み。
    # 的中信頼度スコア = ◎単勝確率ベースの軸信頼度(axis_analysis)を主、gap補正を従
    judgment_score = round(float(np.clip(0.7 * axis_analysis + 0.3 * myomi_axis_score, 0, 100)), 1)

    return {
        'table': table_df, 'axis_analysis': axis_analysis,
        'axis_class': axis_class, 'payout_level_score': payout_level_score,
        'arare_class': arare_class, 'race_trend': race_trend,
        'special_single': special_single, 'fuku_recs': fuku_recs,
        'umatan_recs': umatan_recs, 'umaren_recs': umaren_recs,
        'yuryoku_umatan_pairs': yuryoku_umatan_pairs,
        'wide_recs': wide_recs, 'tan_recs': tan_recs,
        'trio_recs': trio_recs,
        'anaba_umaren': anaba_umaren, 'anaba_umatan': anaba_umatan, 'anaba_wide': anaba_wide,
        'is_invest': is_invest,
        'yuuryoku_tan': yuuryoku_tan, 'yuuryoku_fuku': yuuryoku_fuku,
        'yuuryoku_umatan': yuuryoku_umatan, 'yuuryoku_umaren': yuuryoku_umaren, 'yuuryoku_wide': yuuryoku_wide,
        'myomi_umaren': myomi_umaren, 'myomi_umatan': myomi_umatan,
        # ★ v132/v133: 参考強制フラグ(赤字傾向開催地 または 回収不安定な信頼度区分のレースは
        #   全買い目を備考『参考』に上書き)と、買い目なしレースの補充馬連キー(これも備考『参考』)。
        'force_reference': force_reference,
        # ★v134: 安定運用の単勝/複勝
        'stable_tan_rec': stable_tan_rec, 'stable_tan_tag': stable_tan_tag,
        'stable_fuku_rec': stable_fuku_rec, 'stable_fuku_tag': stable_fuku_tag,
        # ★v135_025: 次点本線枠(複勝安定軸 × 複勝確率50%以上の最上位1頭)
        'subline_rec': subline_rec,
        # ★v135_026: 次点本線枠B(軸wp40+ × 相手複指数24+ × 10頭以下)
        'subline_b_rec': subline_b_rec,
        # ★v135_027: 次点本線枠C(軸wp50+ × 相手複勝40+&騎手50+)馬単・馬連 ※v135_028で停止
        'subline_c_rec': subline_c_rec,
        # ★v135_028: 次点本線枠(勝ち筋マップ)= 本線非採用の sign通過条件(表示専用)
        'jiten_recs': jiten_recs,
        'stat': _stat,   # ★v136_019
        'stat_ura': _stat_ura, 'stat_obs': _stat_obs,   # ★v136_020
        'hl_n': _dq_hl_n, 'hl_m': _dq_hl_m, 'dq_jk': _dq_jk,   # ★v136_021
        'dq_scratched': _dq_scratched, 'race_label': _dq_label,
        'tan_priority_bans': list(tan_priority_bans),
        'conf_tier_precheck': _conf_tier_label,
        'hole_fuku_recs': hole_fuku_recs, 'high_confidence': False,
        'conf_score': 0.0, 'maruren_1pt': None,
        'wide_1pt': None, 'popular_out_high_payout': False,
        'pop_conf_score': 0.0, 'pop_maruren_recs': [],
        'pop_wide_recs': [], 'myomi_axis_score': myomi_axis_score,
        'hole_bans': list(hole_axis_bans), 'fav_bans': list(fav_bans),
        'is_hole_axis': bool(is_hole_axis),
        'is_honmei_axis': is_honmei_axis,
        'axis_ban': axis_ban, 'sub_axis_ban': sub_axis_ban,
        'axis_score': a_score, 'himo_bans': _himo_bans_disp,
        'axis_tan_roi': 0.0, 'axis_tan_n': 0,
        'axis_fuku_roi': 0.0, 'axis_tan_cond': '',
        'judgment_class': judgment_class, 'recommended_action': recommended_action,
        'expected_win_rate': expected_win_rate, 'expected_place_rate': expected_place_rate,
        'judgment_comment': judgment_comment, 'judgment_score': round(judgment_score, 1),
        'axis_zscore': round(single_zscore, 2),
        'axis_win_prob': float(a_winP), 'axis_place_prob': float(a_fukuP),
        # ★v136_008: 新判定の推定値(信頼度の数値に使用。None なら旧式)
        'judge_est_win': (round(judge_est['est_win'], 2) if judge_est else None),
        'judge_est_top3': (round(judge_est['est_top3'], 2) if judge_est else None),
        'judge_quinella': (round(judge_est['quinella'], 2) if judge_est else None),
        'nsl_tan_tier': nsl_tan_tier, 'nsl_fuku_hit': nsl_fuku_hit,
        'spec_axis_type': _spec_axis_type,
        'tan_flags': tan_flags, 'fuku_flags': fuku_flags,
        'gap_conf_factor': gap_conf_factor, 'two_axis_mode': two_axis_mode,
        # ★採用条件: 買い目→発火条件タグ / 買い目→賭け金(円・均等)
        # ★v135_022: 参考枠(表示専用・実弾非関与)
        'reference_bets': reference_bets,
        'gz_umatan_cond': gz_umatan_cond, 'gz_umaren_cond': gz_umaren_cond,
        'gz_wide_cond': gz_wide_cond,          # ★v136_001
        # ★011: 補充馬連(JSON robust条件カスケード)のラベル / 穴印馬番
        'anaba_hidx_bans': anaba_hidx_bans,
        'gz_stakes': gz_stakes,
        'gz_active_conds': _gz_active, 'gz_bankroll': GZ_BANKROLL,
        # ★v135_004: 均等買いメタ(1点単価/レース上限点数/切り捨て点数/総投下額)
        'gz_flat_unit': _GZ_UNIT, 'gz_max_points': _gz_max_pts,
        'gz_dropped_n': gz_dropped_n, 'gz_race_invest': gz_race_invest,
        # ★v135_011: 枠(本線/穴狙い)の内訳
        'gz_active_main': _gz_active_main, 'gz_active_ana': _gz_active_ana,
        'gz_main_points': _n_main, 'gz_ana_points': _n_ana,
        'gz_main_invest': gz_main_invest, 'gz_ana_invest': gz_ana_invest,
        'gz_ana_max_points': _gz_ana_max_pts,
    }


# =====================================================
# ★ 新機能：一括予想（フォルダ指定） v079_一括予想版
# =====================================================
def _run_batch_prediction(app, var_status, var_out_dir, messagebox, filedialog, Path, pd):
    """フォルダ内の全CSVを連続処理して予想ファイルを出力（修正版: スレッド化 + 堅牢化）"""
    folder = filedialog.askdirectory(title='予想対象のCSVファイルが入ったフォルダを選択してください')
    if not folder:
        return

    csv_files = sorted(Path(folder).glob("*.csv"))
    if not csv_files:
        messagebox.showwarning('CSVが見つかりません', '選択したフォルダに *.csv ファイルがありません。')
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

                required_cols = ['単指数', '複指数', '騎手指数', '単勝人気', '単勝オッズ',
                                 '番', 'レース', '展開', '場所', '距離', '出走時刻', '馬名', '騎手']
                missing = [c for c in required_cols if c not in df.columns]
                if missing:
                    raise ValueError(f"列不足: {missing}")

                # ★【バグ修正】'斤量' は required_cols に含まれない任意列なのに
                #   df[numeric_cols] で無条件参照しており、斤量列の無いCSVで
                #   KeyError → そのファイルが丸ごと処理失敗していた。存在列のみ変換する。
                df = to_numeric_racecard(df)
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
<title>地方競馬予想 (グレーゾーン4条件)</title><style>
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
 td.hi-idx{background:#e5ffe5;color:#1e7a1e;font-weight:bold}      /* 単指数90+ */
 td.jk-hot{color:#c0392b;font-weight:bold}                        /* 騎手指数40+ の騎手名 */
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
 td.hi-ana{background:#ffe0b2;color:#a04a00;font-weight:bold}      /* 穴馬の複指数25 */
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
<div class="top"><h1>BADO 地方競馬予想 — 本線重視システム</h1>
<div class="sub">__SRC__ ／ 全__NR__レース ／ ★採用買い目あり __NGZ__レース ／ 均等買い1点__UNIT__円・レース上限__MAXP__点</div>
<div class="legend">
__GZ_LEGEND__
 <span style="color:#ffcf9e">■穴=複指数25×人気4-9×騎手27-62×単指数82内</span>
</div></div>
"""

NL_JOIN_GZ = '\n'


def _h(v):
    """HTMLエスケープ。★v135_010: 馬名/騎手名等に < > & " が入ると表が壊れる
    (XMLとしてパース不能になることを確認済み)。出力前に必ず通す。"""
    import html as _html
    return _html.escape('' if v is None else str(v), quote=True)


def _gz_legend_html():
    """★v135_011: 採用条件レジストリからHTMLレジェンドを生成(枠別に見出しを付ける)。"""
    _parts = []
    for _t in ('main', 'cover'):
        # ★v136_006: 券種ごと(馬単→馬連→ワイド)にまとめ、その中を order 順で並べる。
        _kord = {'umatan': 0, 'umaren': 1, 'wide': 2}
        _items = [(k, v) for k, v in sorted(GZ_COND_DEFS.items(),
                                            key=lambda kv: (_kord.get(kv[1]['kind'], 9), kv[1]['order']))
                  if v['tier'] == _t]
        if not _items:
            continue
        _parts.append(' <b>【%s】</b>' % GZ_TIER_LABEL[_t])
        _parts.extend(
            ' <span style="color:%s">■%s %s %s</span>'
            % (v['bg'], k, GZ_KIND_JP.get(v['kind'], v['kind']), v['short'])
            for k, v in _items)
    # ★v136_018: 次点本線枠(実戦条件シートの観察条件・表示のみ)の説明。
    _sheet_jd = [d for d in (JITEN_MAP_DEFS if JITEN_ENABLE else []) if d.get('src') == 'sheet']
    if _sheet_jd:
        _ax_lab = {'S:ti90max_o19': '軸[単指数90+最大 & 予想オッズ1.9以下]',
                   'S:aw50max_o19': '軸[単勝確率50%+最大 & 予想オッズ1.9以下]'}
        _parts.append(' <b>【次点本線枠(表示のみ)】</b>')
        _parts.extend(
            ' <span style="color:#6b7280">■%s %s %s × %s→紐馬指数1位</span>'
            % (d['tag'], GZ_KIND_JP.get(d['kind'], d['kind']), _ax_lab.get(d['axis'], d['axis']),
               '&'.join(a[2] for a in d['atoms']))
            for d in _sheet_jd)
    # ★v136_003: 参考予想(◎から予想印▲△☆馬への馬連・ワイド)の説明。
    # ★v136_019: 統計予想の説明
    if STAT_DEBA_RACES:
        _parts.append(' <b>【統計予想(参考・実弾外)】</b>')
        _parts.append(' <span style="color:#d9c8f5">■出馬表HTMLの近5走・着別成績・騎手/調教師成績から統計%%を計算し、'
                      '既存の単勝確率と融合(融合%%)。統計印◎○▲△=融合%%の上位。参考買い目=◎-○/◎-▲の馬連・ワイド(確率はHarville推定)。'
                      '検証8/16〜9/13: ◎単勝的中36.3→37.4%%・馬連13.8→15.8%%・ワイド30.0→32.8%%、回収率は70〜80%%で改善なし。'
                      '読込: %s</span>' % _h(stat_summary()))
        _parts.append(' <b>【統計裏付け・観察(記録のみ)】</b>')
        _parts.append(' <span style="color:#d9c8f5">■統計裏付け=本線/次点の軸馬の統計%%が50%%以上(統計印に『裏』、'
                      '買い目に【統計裏付け】)。7/16〜9/13で軸の2着内率が10pt以上高い。'
                      '観察ST1〜ST6=軸を除いた統計%%1位を紐(%s)。買い目の選び方は変えず、'
                      '&lt;出力名&gt;_統計観察記録.csv に記録して10月分で確かめる。</span>'
                      % _h(' / '.join('%s %s %s:%s' % (d['tag'], d['axis'], GZ_KIND_JP.get(d['kind'], d['kind']),
                                                        d['label']) for d in STAT_OBS_DEFS)))
    _parts.append(' <b>【参考予想】</b>')
    _parts.append(
        (' <span style="color:#9aa0a6">■紐 軸=◎(軸馬指数1位) × 予想印○▲△の3頭'
         '(◎との馬連の当たりやすさに対し、配当の妙味を重視した上位3頭)。馬連・ワイド各3点を表示。'
         'いずれも表示のみで実弾対象外。</span>')
        if REF_NEW_LOGIC else
        (' <span style="color:#9aa0a6">■紐 軸=◎(軸馬指数1位) × 予想印▲△☆の馬'
         '(複勝確率%.1f%%以上 or 紐馬指数%.0f以上。該当0頭のときのみ紐馬指数上位3頭)。'
         '馬連・ワイドを表示。いずれも表示のみで実弾対象外。</span>'
         % (HIMO_PLACE_PROB_MIN, HIMO_HIMOBA_IDX_MIN)))
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

    n_gz = sum(1 for _, a, _x in html_races
               if a.get('gz_active_conds') and not a.get('force_reference'))
    parts = [_NAR_HTML_HEAD.replace('<meta charset="utf-8">',
                                    '<meta charset="utf-8">' + bado_version_meta(), 1)
             .replace('__SRC__', _h(src_name))
             .replace('__NR__', str(len(html_races)))
             .replace('__NGZ__', str(n_gz))
             .replace('__GZ_LEGEND__', _gz_legend_html())
             .replace('__UNIT__', str(GZ_FLAT_UNIT))
             .replace('__MAXP__', str(gz_race_max_points()))]
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
            ti = float(row.get('単指数', 0) or 0)
            himo = row.get('紐馬指数', '')
            try:
                ens_v = float(row.get('アンサンブル単勝確率', '') or 0)
                ens_txt = f'{ens_v:.1f}'
            except Exception:
                ens_txt = ''
            # 騎手指数40+で騎手名に色マーク(指数値そのものは表示しない)
            _ki = ex.get('kishu_idx', '')
            try:
                _ki_hot = (float(_ki) >= 40) if _ki != '' else False
            except Exception:
                _ki_hot = False
            jk_cls = ' jk-hot' if _ki_hot else ''
            _is_ana = ('穴' in mk)                      # ★011 穴条件該当馬
            # 各色マーク
            tan_cls = 'hi-tan' if wp >= 40 else 'prob tan'      # 単勝確率40%+
            fuku_cls = 'hi-fuku' if fp >= 70 else 'prob fuku'   # 複勝確率70%+
            idx_cls = ' class="hi-idx"' if ti >= 90 else ''      # 単指数90+
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
                f'<td{idx_cls}>{int(ti) if pd.notna(row.get("単指数")) else ""}</td>'
                f'<td{_ana_cls}>{int(row["複指数"]) if pd.notna(row.get("複指数")) else ""}</td>'
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
            if _jr.get('src') == 'sheet':     # ★v136_018: 実戦条件シートの観察条件
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
            buys.append('<div class="brow" style="opacity:.85">'
                        '<span class="bt" style="background:#9aa0a6">'
                        '%s【参考】</span>'
                        '<b>%s</b> %s &nbsp;%s</div>'
                        % (_kn, _h(_rf['name']), _buy, _note))
        if (not ut and not um and not wd and _st_tan is None and _st_fuku is None
                and not _refs):
            buys.append('<div class="skip">― 買い目なし（本線枠・参考予想とも非該当） ―</div>')
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
        badge = '<span class="badge">★採用買い目あり</span>' if has_gz else ''
        # race_cat(レース分類)部分のみ rmeta に(級はバッジで表示済み)
        _race_cat = rel.split(' / ')[-1] if ' / ' in rel else rel
        parts.append(f"""
        <div class="{'race gz-race' if has_gz else 'race'}">
          <div class="rhead"><span class="rtitle">{_h(venue)} {_h(rno)}R</span>{grade_badge}{conf}
            <span class="rmeta">{_h(dist)}m / 発走{_h(ptime)} / {len(tbl)}頭 / {_h(_race_cat)}</span>{badge}</div>
          <table class="grid"><thead><tr>
            <th>印</th><th>枠</th><th>馬番</th><th>馬名</th><th>性齢</th><th>斤量</th>
            <th>騎手</th><th>{KISHU_DEV_LABEL}</th><th>脚質</th><th>人気</th><th>単オッズ</th>
            <th>単勝率</th><th>複勝率</th><th>期待値</th><th>単指数</th><th>複指数</th><th>軸馬指数</th><th>紐馬指数</th><th>統計印</th><th>統計%</th><th>融合%</th>
          </tr></thead><tbody>{''.join(rows)}</tbody></table>
          <div class="buys">{''.join(buys)}</div>{_data_note_html(analysis)}{_stat_block_html(analysis)}</div>""")
    parts.append('<div class="footer">CI下限&lt;100%のため利益は統計的に未確定。損失許容の範囲で。</div></body></html>')
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
NOTE_DISCLAIMER = ('掲載している確率・指数・買い目は、過去データから作成したAIモデルによる予測です。'
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
        force_ref = bool(analysis.get('force_reference', False))
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
            ti = _f(row.get('単指数'))
            fi = _f(row.get('複指数'))
            od = _f(row.get('単勝オッズ'), 0.0)
            ens = _f(row.get('アンサンブル単勝確率'))
            himo = _f(row.get('紐馬指数'))
            kdev = _f(ex.get('kishu_dev', ''))
            kidx = _f(ex.get('kishu_idx', ''))
            pop = _f(row.get('単勝人気'))
            pop_s = str(int(pop)) if (pop is not None and pop > 0) else '-'
            style = str(row.get('展開', '') or '').strip() or '-'
            jk = str(row.get('騎手', '') or '')
            jk_html = ('<span class="jkhot">%s</span>' % _h(jk)) if (kidx is not None and kidx >= 40) else _h(jk)
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
                + '<td class="g1">%s</td>' % _note_bar(wp, 'w' + (' hi' if wp >= 40 else ''), 60.0)
                + '<td>%s</td>' % _note_bar(fp, 'p' + (' hi' if fp >= 70 else ''), 100.0)
                + '<td class="num%s">%.0f</td>' % (' hot' if ev >= 110 else '', ev)
                + '<td class="num g1%s">%s</td>' % (' hot' if (ti or 0) >= 90 else '',
                                                    ('%d' % ti) if ti is not None else '-')
                + '<td class="num%s">%s</td>' % (' anav' if is_ana else '',
                                                 ('%d' % fi) if fi is not None else '-')
                + '<td class="num">%s</td>' % (('%.1f' % ens) if ens is not None else '-')
                + '<td class="num">%s</td>' % (('%.0f' % himo) if himo is not None else '-')
                + '<td class="num%s">%s</td>' % (' hot' if (kdev or 0) >= 60 else '',
                                                 ('%.1f' % kdev) if kdev is not None else '-')
                + '</tr>')

        # ── 推奨買い目(本線枠) ──
        main_tks = []
        stakes = analysis.get('gz_stakes', {})
        for kind, recs_key, cond_key, arrow in (('wide', 'wide_recs', 'gz_wide_cond', '-'),
                                                ('umaren', 'umaren_recs', 'gz_umaren_cond', '-'),
                                                ('umatan', 'umatan_recs', 'gz_umatan_cond', '→')):
            cmap = analysis.get(cond_key, {}) or {}
            for rec in analysis.get(recs_key, []) or []:
                cn = cmap.get((rec.ban1, rec.ban2), '')
                if gz_tier(cn) != 'main':
                    continue
                main_tks.append((kind, rec.ban1, rec.ban2, arrow, cn))
        has_main = bool(main_tks) and not force_ref
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
            _lab = ('推奨買い目' if not force_ref
                    else '推奨買い目（この開催は見送り推奨のため参考表示）')
            _items = ''.join(
                _tk(GZ_KIND_JP.get(k, k), '%d%s%d' % (b1, ar, b2),
                    '%s・%s' % (names_by_ban.get(b1, ''), names_by_ban.get(b2, '')),
                    (cn if NOTE_SHOW_COND_CODE else ''))
                for k, b1, b2, ar, cn in main_tks)
            buys.append('<div class="bgrp %s"><div class="bl">%s</div><div class="tks">%s</div></div>'
                        % ('main' if not force_ref else 'ref', _lab, _items))

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
        if NOTE_SHOW_REFERENCE:
            ref_items = []
            for rf in analysis.get('reference_bets') or []:
                kn = GZ_KIND_JP.get(rf.get('kind'), rf.get('kind'))
                sep = GZ_KIND_SEP.get(rf.get('kind'), '-')
                combo = ' ／ '.join('%d%s%d' % (rf['axis_ban'], sep, b) for b in rf.get('partner_bans', []))
                if combo:
                    ref_items.append(_tk(kn, combo, '◎から印の馬へ'))
            if ref_items:
                buys.append('<div class="bgrp ref"><div class="bl">参考買い目</div>'
                            '<div class="tks">%s</div></div>' % ''.join(ref_items))
        if not buys:
            buys.append('<div class="none">このレースは買い目なし（見送り）</div>')

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
        flag = '<span class="flag">推奨買い目あり</span>' if has_main else ''
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
              '<th colspan="3" class="g1">AI予測</th><th colspan="5" class="g1">指数</th></tr>'
              '<tr><th class="g1">人気</th><th>単勝</th><th class="g1">勝率%</th><th>複勝率%</th><th>期待値</th>'
              '<th class="g1">単指数</th><th>複指数</th><th>軸馬</th><th>紐馬</th><th>騎手<br>補正</th></tr></thead><tbody>'
            + ''.join(trs) + '</tbody></table></div>'
            + '<div class="buys">%s</div></section>' % ''.join(buys))

    # ── 本日の推奨買い目(発走順) ──
    index_rows.sort(key=lambda t: (t[0] or '99:99', str(t[1]), t[2]))
    if index_rows:
        _ir = ''.join(
            '<tr><td class="t num">%s</td><td class="race"><a href="#%s">%s %dR</a></td>'
            '<td><span class="chips">%s</span></td><td class="ax">◎ %s</td></tr>'
            % (_h(t), a, _h(v), r,
               ''.join('<span class="chip"><span class="ty">%s</span><b>%s</b></span>'
                       % (_h(k), _h(c)) for k, c in tks), _h(n))
            for t, v, r, a, tks, n in index_rows)
        index_html = ('<div class="index"><table><thead><tr><th>発走</th><th>レース</th><th>推奨買い目</th>'
                      '<th class="ax">本命◎</th></tr></thead><tbody>%s</tbody></table></div>' % _ir)
    else:
        index_html = '<div class="empty">本日の推奨買い目はありません（全レース見送り）</div>'

    guide_html = (
        '<div class="guide"><dl>'
        '<dt>印</dt><dd><span class="marks">'
        + ' '.join('%s %s' % (_note_mark(k), v) for k, v in
                   ((('◎', '本命'), ('○', '対抗'), ('▲', '単穴'), ('△', '連下'), ('穴', '穴馬')) if REF_NEW_LOGIC
                    else (('◎', '本命'), ('○', '対抗'), ('▲', '単穴'), ('△', '連下'), ('☆', '押さえ'), ('穴', '穴馬'))))
        + '</span></dd>'
        '<dt>脚質</dt><dd>逃=逃げ・先=先行・差=差し・追=追込。</dd>'
        '<dt>騎手の色</dt><dd>騎手名が<span class="jkhot">色付き</span>は、騎手指数が高い(40以上)騎手です。</dd>'
        '<dt>勝率・複勝率</dt><dd>AIが推定した「1着になる確率」「3着以内に入る確率」(%)。勝率40%以上・複勝率70%以上は濃い色で表示。</dd>'
        '<dt>期待値</dt><dd>勝率×単勝オッズ。100を超えるほど、オッズに対して割安な馬です。</dd>'
        '<dt>単指数</dt><dd>1着を争う能力の指数。</dd>'
        '<dt>複指数</dt><dd>3着以内に来る安定度の指数。<span class="anav">橙色</span>は穴候補の馬。</dd>'
        '<dt>軸馬</dt><dd>軸馬指数。AIの勝率と単勝オッズを合わせた「軸としての強さ」(%)。◎はこの1位です。</dd>'
        + ('<dt>○▲△</dt><dd>◎との組み合わせで、当たりやすさに加えて「配当の妙味」を重視して選んだ3頭。参考買い目(◎から馬連・ワイド各3点)の相手です。</dd>' if REF_NEW_LOGIC else '') +
        '<dt>紐馬</dt><dd>紐馬指数(0〜100)。複指数・騎手・展開などから見た「相手としての来やすさ」。推奨買い目の相手選びに使います。</dd>'
        '<dt>騎手補正</dt><dd>騎手の腕を偏差値で表したもの(50が平均)。60以上は好騎手。</dd>'
        '<dt>緑の太字</dt><dd>注目の値(期待値110以上・単指数90以上・騎手補正60以上)。</dd>'
        + ('<dt>軸級 S〜D</dt><dd>◎が1着になる推定確率の区分。%s / D=それ未満。</dd>'
           '<dt>信頼度</dt><dd>◎が3着以内に入る推定確率(%%)。%s。</dd>'
           % (' / '.join('%s=%g%%以上' % (_g, _c) for _g, _c in JUDGE_GRADE_CUTS),
              '・'.join('「%s」%g以上' % (_lb, _mn) for _mn, _lb, *_ in JUDGE_CONF_TIERS if _mn >= 0))
           if USE_NEW_JUDGE else
           '<dt>信頼度</dt><dd>S〜Dは軸馬の信頼度(Sが最も高い)。数字は0〜100で、そのレースの買い目が的中しやすいかの目安です。</dd>') +
        '<dt>推奨買い目</dt><dd>過去データで条件を検証したBADOの本命買い目です（黄色の枠のレース）。組番は馬番です。</dd>'
        '<dt>参考買い目</dt><dd>◎から印の馬への組み合わせ。推奨買い目より点数が多く、参考としてご覧ください。</dd>'
        '</dl></div>')

    title = '%s 地方競馬予想%s' % (NOTE_BRAND, (' ' + date_label) if date_label else '')
    out = ['<!DOCTYPE html><html lang="ja"><head><meta charset="utf-8">' + bado_version_meta() +
           '<meta name="viewport" content="width=device-width,initial-scale=1">'
           '<title>%s</title><style>%s</style></head><body><div class="wrap">' % (_h(title), _NOTE_CSS),
           '<header class="cover"><div class="brand">%s ／ 地方競馬AI予想</div>' % _h(NOTE_BRAND),
           '<h1>%s</h1>' % _h(date_label or '地方競馬予想'),
           '<div class="venues">%s</div>' % _h('・'.join(str(v) for v in venues)),
           '<div class="kpis"><div class="kpi"><b>%d</b><span>予想レース</span></div>'
           '<div class="kpi"><b>%d</b><span>推奨買い目のあるレース</span></div>'
           '<div class="kpi"><b>%d</b><span>推奨買い目（点）</span></div></div></header>'
           % (len(cards), n_pick_races, n_pick_pts),
           '<h2>本日の推奨買い目（発走順）</h2>', index_html,
           '<h2>予想表の見方</h2>', guide_html,
           '<h2>全レースの予想</h2>']
    out.extend(cards)
    out.append('<div class="foot">%s<div class="ver">%s %s</div></div></div></body></html>'
               % (_h(NOTE_DISCLAIMER), _h(NOTE_BRAND), _h(BADO_VERSION)))
    return '\n'.join(out)


# ──────────────────────────────────────────────────────────────
# Excel / CSV 出力（v077完全版）
# ──────────────────────────────────────────────────────────────
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
            ws.cell(row=current_row, column=1, value="AI BADO 地方競馬予想 統計強化 実績的中率学習テーブル")
            ws.cell(row=current_row, column=1).fill = title_fill
            ws.cell(row=current_row, column=1).font = title_font
            ws.cell(row=current_row, column=1).alignment = Alignment(horizontal='center')
            current_row += 1

            ws.cell(row=current_row, column=1, value=" 緑=高確率/高EV | 黄=期待値110+ | 青=強軸 | オレンジ=穴馬軸 | 淡い緑=単指数89以上 | 淡い青=単勝オッズ9.9以下 | 淡い藍=BADO指数65以上(強軸) | 淡い橙=紐馬指数65以上(強紐馬) | ★強い青=軸の分析値75以上 | ★緑=A軸以上(判定) | 馬番【緑】=軸候補(BADO80+/複23-25&人気4+&BADO60+/単90+&1番人気&複勝80+) | 馬番【薄ピンク】=紐候補(紐馬指数60+&複勝40+ または BADO指数2位) | 馬名【黄色】=単指数≥90 & 人気1位 & 複勝確率≥80 の鉄板馬 | 馬名【ピンク】=鉄板馬が存在するレースのみ：紐馬指数1位 or 複指数1位 or BADO指数1位 or 単勝確率1位 の馬 ")
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
                        _extras[_b] = dict(
                            waku=_to_int_z2h(_gr.get('枠')) if '枠' in group.columns else None,
                            seirei=str(_gr.get('性齢', '') or ''),
                            kinryo=(str(_gr.get('斤量', '') or '')
                                    if '斤量' in group.columns else ''),
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
                    is_iron_horse = '単90以上鉄板馬' in str(analysis.get('axis_class', ''))

                    # ===== 新規追加: 鉄板馬（単指数≥90 & 人気1位 & 複勝確率≥80）の検出と、該当レースのみのトップ指標馬特定 =====
                    iron_mask = (
                        (analysis['table']['単指数'] >= 90) &
                        (analysis['table']['単勝人気'] == 1) &
                        (analysis['table']['複勝確率'] >= 80)
                    )
                    has_iron = bool(iron_mask.any())
                    iron_bans = set(int(b) for b in analysis['table'].loc[iron_mask, '番'].dropna().astype(int).tolist()) if has_iron else set()

                    # ========== Excel書き込み ==========
                    race_title = f"{name[0]} {name[1]}m {name[2]}R ({name[3]}) | 荒れ度:{analysis['payout_level_score']} | {analysis['arare_class']}"
                    ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=14)
                    cell = ws.cell(row=current_row, column=1, value=race_title)
                    cell.fill = header_fill
                    cell.font = header_font
                    current_row += 1

                    headers = ['番','馬名','騎手',KISHU_DEV_LABEL,'展開','単指数','複指数','BADO指数','紐馬指数','印','単勝人気','単勝オッズ','単勝確率','複勝確率','単勝期待値','統計印','統計勝率','融合勝率']
                    for col_idx, h in enumerate(headers, 1):
                        c = ws.cell(row=current_row, column=col_idx, value=h)
                        c.fill = header_fill
                        c.font = header_font
                        c.border = thin_border
                        c.alignment = Alignment(horizontal='center')
                    current_row += 1

                    for _, row in analysis['table'].iterrows():
                        pop = int(row['単勝人気']) if not pd.isna(row['単勝人気']) else 99
                        fuku_prob = round(row['複勝確率'],1)
                        single_idx = int(row['単指数']) if not pd.isna(row['単指数']) else 0
                        win_odds = round(row['単勝オッズ'],1)
                        # ===== 軸候補（緑）/紐候補（青）判定用の値を取得 =====
                        bado_val  = float(row['BADO指数'])  if not pd.isna(row['BADO指数'])  else 0.0
                        fuku_idx  = float(row['複指数'])     if not pd.isna(row['複指数'])     else 0.0
                        tan_idx   = float(row['単指数'])     if not pd.isna(row['単指数'])     else 0.0
                        # 軸候補（緑）: BADO80+ / (複指数23〜25 + 人気4番以降 + BADO60+) / (単指数90+ + 単勝1番人気 + 複勝確率80+)
                        is_green_axis = (
                            (bado_val >= 80) or
                            (23 <= fuku_idx <= 25 and pop >= 4 and bado_val >= 60) or
                            (tan_idx >= 90 and pop == 1 and fuku_prob >= 80)
                        )
                        # 紐候補（青）: (紐馬指数60+ かつ 複勝確率40+) または BADO指数2位
                        ban_no = int(row['番']) if not pd.isna(row['番']) else None
                        # ===== 新規: 鉄板馬（黄色） / トップ指標馬（ピンク・鉄板馬存在時のみ） =====
                        is_iron_horse_row = ban_no in iron_bans if has_iron else False
                        vals = [
                            int(row['番']) if not pd.isna(row['番']) else '',
                            row['馬名'], row['騎手'],
                            # ★v135_005: 表示専用 補正騎手指数(偏差値)
                            (float(row[KISHU_DEV_COL])
                             if KISHU_DEV_COL in row.index
                             and not pd.isna(row[KISHU_DEV_COL]) else ''),
                            str(row['展開']),
                            single_idx,
                            int(row['複指数']) if not pd.isna(row['複指数']) else '',
                            row['BADO指数'], row['紐馬指数'], row['印'],
                            pop,
                            win_odds,
                            round(row['単勝確率'],1),
                            fuku_prob,
                            int(row['単勝期待値']) if not pd.isna(row['単勝期待値']) else '',
                            # ★v136_019 統計予想
                            str(row.get('統計印', '') or ''),
                            (round(float(row['統計勝率']), 1) if pd.notna(row.get('統計勝率')) else ''),
                            (round(float(row['融合勝率']), 1) if pd.notna(row.get('融合勝率')) else ''),
                        ]
                        for col_idx, v in enumerate(vals, 1):
                            c = ws.cell(row=current_row, column=col_idx, value=v)
                            c.border = thin_border
                            c.alignment = Alignment(horizontal='center')

                            # ===== 馬番セルに軸候補（緑）マーク（v099改: 紐候補の薄ピンクは廃止） =====
                            if col_idx == 1:
                                if is_green_axis:
                                    c.fill = PatternFill(start_color="00B050", end_color="00B050", fill_type="solid")
                                    c.font = Font(bold=True, color="FFFFFF")

                            # ===== 馬名セル（col=2）は黄色（鉄板馬）のみ（v099改: ピンク=トップ指標馬は廃止） =====
                            if col_idx == 2:  # 馬名
                                if is_iron_horse_row:
                                    c.fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
                                    c.font = Font(bold=True, color="000000")

                            if col_idx == 11 and 1 <= pop <= 4:
                                c.fill = PatternFill(start_color="CFE2F3", end_color="CFE2F3", fill_type="solid")

                            if col_idx == 14 and fuku_prob >= 60:
                                c.fill = PatternFill(start_color="B6D7A8", end_color="B6D7A8", fill_type="solid")

                            if col_idx == 10 and v == '◎':
                                c.fill = iron_red
                                c.font = Font(bold=True)
                            elif col_idx == 10 and '穴' in str(v):   # ★011 穴条件該当馬
                                c.fill = hole_orange
                                c.font = Font(bold=True, color="A04A00")
                            elif col_idx == 13 and isinstance(v, (int,float)) and v >= 40:
                                c.fill = good_green
                            elif col_idx == 15 and isinstance(v, (int,float)) and v >= 100:  # v099改: 単勝期待値100+は無条件
                                c.fill = high_yellow
                            elif col_idx == 7 and isinstance(v, (int,float)) and v >= 23:  # v099改: 複指数23+は無条件
                                c.fill = hole_orange

                            if col_idx == 6 and isinstance(v, (int, float)) and v >= 89:
                                c.fill = high_single_fill

                            if col_idx == 12 and isinstance(v, (int, float)) and v <= 9.9:
                                c.fill = low_odds_fill

                            if col_idx == 8 and isinstance(v, (int, float)) and v >= 65:
                                c.fill = new_s_axis_fill

                            if col_idx == 9 and isinstance(v, (int, float)) and v >= 65:
                                c.fill = himba_idx_fill
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
                    for _tk, _tnote in (('main', '本線枠｜実戦条件シート(紐条件探索3・4、7/16〜9/23)より選定・馬連2/ワイド1'),
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

                    # ★v136_020: 統計裏付け・観察(統計紐)の行(表示・記録のみ)
                    for _sl in _dq_text_lines(analysis) + _stat_text_lines(analysis):   # ★v136_021
                        _sc = ws.cell(row=current_row, column=1, value=_sl)
                        _sc.font = Font(bold=True, size=10, color="5B3F8F")
                        current_row += 1

                    current_row += 1

                    # ========== CSV出力 ==========
                    title = f"{name[0]}  {name[1]}m {name[2]}R ({name[3]}),,,,,,,,,,,, "
                    f.write(title + '\n')

                    f.write('馬番,馬名,騎手,' + KISHU_DEV_LABEL + ',展開,単指数,複指数,BADO指数,紐馬指数,印,単勝人気,単勝オッズ,単勝確率,複勝確率,単勝期待値,統計印,統計勝率,融合勝率\n')

                    unfold_map = {'逃': '逃げ', '先': '先行', '差': '差し', '追': '追込', '中': '中団', '捲': '捲り', 'マ': 'マーク'}
                    for _, row in analysis['table'].iterrows():
                        unfold_str = str(row['展開'])
                        for key, val in unfold_map.items():
                            if key in unfold_str:
                                unfold_str = unfold_str.replace(key, val)
                        line = '{},{},{},{},{},{},{},{},{},{},{},{:.1f},{:.1f},{:.1f},{},{},{},{}\n'.format(
                            int(row['番']) if not pd.isna(row['番']) else '',
                            row['馬名'], row['騎手'],
                            # ★v135_005: 表示専用 補正騎手指数(偏差値)
                            (row[KISHU_DEV_COL] if KISHU_DEV_COL in row.index else ''),
                            unfold_str,
                            int(row['単指数']) if not pd.isna(row['単指数']) else '',
                            int(row['複指数']) if not pd.isna(row['複指数']) else '',
                            row['BADO指数'], row['紐馬指数'], row['印'],
                            int(row['単勝人気']) if not pd.isna(row['単勝人気']) else '',
                            row['単勝オッズ'],
                            row['単勝確率'], row['複勝確率'], row['単勝期待値'],
                            str(row.get('統計印', '') or ''),
                            ('%.1f' % float(row['統計勝率']) if pd.notna(row.get('統計勝率')) else ''),
                            ('%.1f' % float(row['融合勝率']) if pd.notna(row.get('融合勝率')) else ''),
                        )
                        f.write(line)

                    f.write(',,,,,,,,,,,,,\n')
                    f.write('「レース分析」,,,,,,,,,,,,,\n')

                    f.write(f'軸馬: {axis_ban_str}番 {axis_name_str},,,,,,,,,,,,,\n')

                    is_iron_horse = '単90以上鉄板馬' in str(analysis.get('axis_class', ''))
                    axis_line = f'軸の分析値: {analysis["axis_analysis"]}/100  妙味軸度: {analysis["myomi_axis_score"]}点  軸z-score: {analysis.get("axis_zscore", 0):.2f}'
                    if is_iron_horse:
                        axis_line += ' ★単90以上鉄板馬'
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
                        f.write('おすすめ買い目 (回収率90%超マージンを見込んで):,,,,,,,,,,,,,\n')

                    f.write(',,,,,,,,,,,,,\n')
                    f.write('買い目,,,,,,,,,,,,,\n')

                    is_iron_horse = '単90以上鉄板馬' in str(analysis.get('axis_class', ''))
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
    var_status     = tk.StringVar(value='CSVファイルを開いてください(Ctrl+O)')
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
                                        ('回収率重視 λ=%g' % REF_VALUE_LAMBDA) if REF_NEW_LOGIC else '旧')
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
    btn_csv = _button(dc, 'CSVを開く  (Ctrl+O)', lambda: _open_csv(), kind='brand')
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
    ttk.Checkbutton(_flt, text='★採用のみ', variable=var_star_only, style='Bado.TCheckbutton',
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
        ('ti',   '単指数',    58,  'center', '単指数'),
        ('fi',   '複指数',    58,  'center', '複指数'),
        ('ens',  '軸馬指数',  70,  'e',      'アンサンブル単勝確率'),
        ('himo', '紐馬指数',  70,  'e',      '紐馬指数'),
        ('smk',  '統計印',    52,  'center', None),        # ★v136_019
        ('sp',   '統計%',     60,  'e',      '統計勝率'),
        ('fu',   '融合%',     60,  'e',      '融合勝率'),
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
    for _i, _k in enumerate(('勝率', '複勝率', '単指数', '人気', 'オッズ')):
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
    tv_sum.column('#0', width=_colw(170, '券種', '▾ 次点本線枠（表示のみ）'), anchor='w', stretch=False)
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
            title='レースデータCSVを選択', filetypes=[('CSV files', '*.csv'), ('All files', '*.*')])
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
        if analysis.get('force_reference'):
            return False
        return any(v > 0 for v in (analysis.get('gz_stakes') or {}).values())

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
            required_cols = ['単指数', '複指数', '騎手指数', '単勝人気', '単勝オッズ', '番', 'レース', '展開', '場所', '距離', '出走時刻', '馬名', '騎手']
            missing = [c for c in required_cols if c not in df.columns]
            if missing:
                messagebox.showerror('列不足', f'必要な列がありません: {missing}')
                _set_status('列不足のため読み込めませんでした')
                return
            df = to_numeric_racecard(df)
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
        lbl_race_meta.config(text=f'{dist}m ・ 発走 {tm} ・ {n_heads}頭')
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
        if _has_star(a):
            _chip(chip_row, '★ 採用買い目あり', P['GZ']).pack(side='left', padx=(px(8), 0))
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
                _fmt(row.get('単指数'), 0, ''),
                _fmt(row.get('複指数'), 0, ''),
                _fmt(row.get('アンサンブル単勝確率'), 1),
                _fmt(row.get('紐馬指数'), 1),
                str(row.get('統計印', '') or ''),
                _fmt(row.get('統計勝率'), 1, ''),
                _fmt(row.get('融合勝率'), 1, ''),
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
            _stat_lbls['単指数'].config(text=_fmt(r0.get('単指数'), 0))
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
                                      f"融合%: {_stc['p_fused'][_x['ban1']]:.1f} / {_stc['p_fused'][_x['ban2']]:.1f}"),
                              tags=('normal',))
            any_row = True
        # ★v136_020: 統計裏付け・観察(統計紐・記録のみ)
        _ura = a.get('stat_ura') or {}
        _obs_ok = [o for o in (a.get('stat_obs') or []) if o['ok']]
        if _ura or _obs_ok:
            sec = tv_sum.insert('', 'end', text='統計裏付け・観察（記録のみ）', open=True, tags=('section',),
                                values=('', '', '', '', '軸の統計勝率50%以上=統計裏付け'))
            _seen = {}
            for _k in ('UR1', 'UR2', 'WD1', 'OU1'):
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
            sec = tv_sum.insert('', 'end', text='参考予想 ◎→○▲△', open=True, tags=('section',),
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
            tv_sum.insert('', 'end', text='買い目なし', values=('', '', '', '', 'このレースは見送り'),
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
            messagebox.showwarning('ファイル未選択', 'CSVファイルを選択してください。')
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