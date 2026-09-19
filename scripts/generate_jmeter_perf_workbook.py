# -*- coding: utf-8 -*-
"""Generate the JMeter overall performance-test Excel workbook."""
from __future__ import annotations

from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.formatting.rule import FormulaRule
from openpyxl.worksheet.datavalidation import DataValidation

OUT = Path(__file__).resolve().parents[1] / "docs" / "JMeter_全体パフォーマンステスト完全解説.xlsx"

NAVY = "1B365D"
TEAL = "0F6C8C"
GOLD = "C4A35A"
RED = "9B2C2C"
GREEN = "276749"
ORANGE = "C05621"
SLATE = "2D3748"
PALE = "F7FAFC"
PALE2 = "EDF2F7"
PALE_TEAL = "E6F3F7"
PALE_GOLD = "F8F1DE"
PALE_RED = "FED7D7"
PALE_GREEN = "C6F6D5"
WHITE = "FFFFFF"
BLACK = "1A202C"

THIN = Border(
    left=Side(style="thin", color="CBD5E0"),
    right=Side(style="thin", color="CBD5E0"),
    top=Side(style="thin", color="CBD5E0"),
    bottom=Side(style="thin", color="CBD5E0"),
)
WRAP = Alignment(wrap_text=True, vertical="top")
WRAP_C = Alignment(wrap_text=True, vertical="center")
CENTER = Alignment(wrap_text=True, vertical="center", horizontal="center")


def fill(hex_color: str) -> PatternFill:
    return PatternFill("solid", fgColor=hex_color)


def font(bold=False, color=WHITE, size=11, name="Calibri") -> Font:
    return Font(name=name, bold=bold, color=color, size=size)


def apply_widths(ws: Worksheet, widths: dict[int, float]) -> None:
    for col, width in widths.items():
        ws.column_dimensions[get_column_letter(col)].width = width


def banner(ws: Worksheet, row: int, cols: int, text: str, color: str = NAVY, size: int = 16) -> None:
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=cols)
    cell = ws.cell(row, 1, text)
    cell.fill = fill(color)
    cell.font = font(True, WHITE, size)
    cell.alignment = Alignment(vertical="center", wrap_text=True, indent=1)
    ws.row_dimensions[row].height = 28 if size < 18 else 36


def note(ws: Worksheet, row: int, cols: int, text: str, color: str = PALE_GOLD) -> None:
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=cols)
    cell = ws.cell(row, 1, text)
    cell.fill = fill(color)
    cell.font = font(False, SLATE, 10)
    cell.alignment = Alignment(vertical="center", wrap_text=True, indent=1)
    ws.row_dimensions[row].height = 36


def headers(ws: Worksheet, row: int, titles: list[str], color: str = TEAL) -> None:
    for i, title in enumerate(titles, 1):
        cell = ws.cell(row, i, title)
        cell.fill = fill(color)
        cell.font = font(True, WHITE, 11)
        cell.alignment = CENTER
        cell.border = THIN
    ws.row_dimensions[row].height = 24
    if ws.freeze_panes is None:
        ws.freeze_panes = f"A{row + 1}"


def put_rows(
    ws: Worksheet,
    start: int,
    rows: list[list[object]],
    stripe: bool = True,
    filter_header: int | None = None,
    apply_filter: bool = True,
) -> int:
    r = start
    ncols = 1
    for i, row in enumerate(rows):
        bg = PALE if (stripe and i % 2 == 0) else WHITE
        max_len = 0
        ncols = max(ncols, len(row))
        for c, val in enumerate(row, 1):
            cell = ws.cell(r, c, val)
            cell.fill = fill(bg)
            cell.font = font(False, BLACK, 10)
            cell.alignment = WRAP
            cell.border = THIN
            if val is not None:
                max_len = max(max_len, str(val).count("\n"))
        ws.row_dimensions[r].height = min(120, 18 + max(18, max_len * 14))
        r += 1
    if rows and apply_filter:
        top = filter_header if filter_header is not None else start - 1
        ws.auto_filter.ref = f"A{top}:{get_column_letter(ncols)}{r - 1}"
    return r


def section_title(ws: Worksheet, row: int, cols: int, text: str, color: str = GOLD) -> int:
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=cols)
    cell = ws.cell(row, 1, text)
    cell.fill = fill(color)
    cell.font = font(True, NAVY, 12)
    cell.alignment = Alignment(vertical="center", wrap_text=True, indent=1)
    ws.row_dimensions[row].height = 22
    return row + 1


def sheet_base(wb: Workbook, name: str, tab_color: str = TEAL) -> Worksheet:
    ws = wb.create_sheet(name)
    ws.sheet_properties.tabColor = tab_color
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.page_setup.paperSize = ws.PAPERSIZE_A3
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_title_rows = "1:3"
    ws.page_setup.horizontalCentered = True
    ws.oddHeader.left.text = "JMeter 全体パフォーマンステスト完全解説"
    ws.oddFooter.left.text = name
    ws.oddFooter.right.text = "Page &P / &N"
    return ws


def build() -> None:
    wb = Workbook()
    # remove default later
    default = wb.active
    default.title = "_tmp"

    cover(wb)
    toc(wb)
    principles(wb)
    history(wb)
    glossary(wb)
    test_types(wb)
    viewpoints(wb)
    reports(wb)
    cautions(wb)
    injection(wb)
    apigw(wb)
    alb(wb)
    ecs(wb)
    aurora(wb)
    rdsproxy(wb)
    elasticache(wb)
    report_server(wb)
    correlation(wb)
    metrics(wb)
    passfail(wb)
    failures(wb)
    checklist(wb)
    analogies(wb)
    commands(wb)
    references(wb)

    wb.remove(default)
    wb.properties.title = "JMeter全体パフォーマンステスト完全解説"
    wb.properties.creator = "JMeter Performance Test Guide"
    wb.properties.subject = "JMeter / AWS ECS / Aurora Serverless v2 / RDS Proxy / ElastiCache / 帳票"
    wb.properties.description = (
        "JMeter 5.6.3 を用いた全体パフォーマンステストの観点・レポート・注意点・"
        "用語・API Gateway/ALB/ECS/Aurora Serverless v2/RDS Proxy/ElastiCache/帳票サーバの測り方"
    )
    OUT.parent.mkdir(parents=True, exist_ok=True)
    wb.save(OUT)
    print(f"wrote {OUT}")


def cover(wb: Workbook) -> None:
    ws = sheet_base(wb, "00_表紙", NAVY)
    apply_widths(ws, {1: 28, 2: 88, 3: 28, 4: 28})
    banner(ws, 1, 4, "JMeter を用いた全体パフォーマンステスト完全解説", NAVY, 20)
    ws.row_dimensions[1].height = 42
    note(
        ws,
        2,
        4,
        "小学生にも分かるたとえ話と、動作原理・歴史・最新仕様（2026-09）を一枚にまとめた実務ブック。"
        "詳細な文章は同梱の Markdown を、試験当日の表とチェックは本 Excel を使う。",
        PALE_TEAL,
    )
    headers(ws, 4, ["項目", "内容", "値・版", "備考"])
    put_rows(
        ws,
        5,
        [
            ["文書の役割", "全体性能試験の設計・実行・報告の共通言語", "ガイド＋チェックリスト", "計画書の雛形に転記可"],
            ["対象ツール", "Apache JMeter", "5.6.3（公式最新安定、2024-01-07）", "Java 8+ / 実行は Java 17 推奨"],
            ["本試験の実行", "非 GUI（CLI）", "jmeter -n -t ... -l ... -e -o ...", "GUI は作成とデバッグ専用"],
            ["標準の打ち込み位置", "アプリ能力は ALB 直下から", "API Gateway は別ラン", "門を測るか厨房を測るかを混ぜない"],
            ["ECS の測り方", "先に 1 タスクの飽和点", "その後 2/4/8 で直線性", "サービス平均 CPU だけでは不合格"],
            ["Aurora Serverless v2", "ACU 固定ランと伸縮ランを分ける", "試験中はオートポーズ禁止", "伸縮待ちをアプリ遅延と呼ばない"],
            ["RDS Proxy", "ピン留め率を必須観測", "16KB 超 SQL は全エンジンでピン", "ピン留め多数なら多重化が死んでいる"],
            ["ElastiCache", "ホットとコールドを別ラン", "EngineCPU とヒット率", "ヒット 100% はキャッシュ単体試験"],
            ["帳票サーバ", "あり/なしのミックスを分離", "同期は 29 秒の壁", "ECS 増設では印刷室は増えない"],
            ["一次資料", "JTL + ログ + CloudWatch", "HTML は説明用", "平均より p95/p99"],
            ["情報時点", "2026-09", "クォータはアカウントで確認", "Service Quotas を試験計画に写す"],
        ],
        apply_filter=True,
        filter_header=4,
    )
    r = section_title(ws, 17, 4, "一枚絵（小学生向け）", GOLD)
    put_rows(
        ws,
        r,
        [
            ["お客役", "JMeter の仮想ユーザ（スレッド）", "同時に何人来るか", "人間より高頻度になりがち"],
            ["校門の警備", "API Gateway", "1秒の人数制限・学生証", "ここを通すと門の試験になる"],
            ["配膳口", "ALB", "空いている窓へ案内", "アプリ試験の標準エントリ"],
            ["調理員ひとり", "ECS タスク", "実際に料理を作る単位", "先に一人分の皿数を測る"],
            ["手元の棚", "ElastiCache", "よく出すおかず", "空だと全員が倉庫へ走る"],
            ["整理券機", "RDS Proxy", "倉庫の席の使い回し", "専用ロッカー（ピン留め）で死ぬ"],
            ["伸縮する倉庫", "Aurora Serverless v2", "床（ACU）が伸び縮み", "広げる時間は別に測る"],
            ["印刷室", "帳票サーバ", "成績表を丁寧に印刷", "同時台数が天井"],
        ],
        stripe=True,
        apply_filter=False,
    )
    note(
        ws,
        27,
        4,
        "約束: 厨房の速さを知りたいのに校門の幅だけを測ってはいけない。測りたい部屋の直前からお客を入れる。",
        PALE_RED,
    )


def toc(wb: Workbook) -> None:
    ws = sheet_base(wb, "01_目次", TEAL)
    apply_widths(ws, {1: 10, 2: 32, 3: 70, 4: 36})
    banner(ws, 1, 4, "目次と読み方", NAVY, 16)
    note(ws, 2, 4, "試験設計は 05〜09、当日オペは 20 チェックリスト、障害切り分けは 19 失敗の見分け。文章の深掘りは Markdown。")
    headers(ws, 4, ["シート", "名前", "中身", "いつ使う"])
    rows = [
        ["00", "表紙", "対象版、一枚絵、約束", "最初に共有"],
        ["01", "目次", "この表", "ナビ"],
        ["02", "原則", "測る順番とやってはいけないこと", "キックオフ"],
        ["03", "歴史年表", "なぜその層とツールが生まれたか", "説明・教育"],
        ["04", "用語辞典", "JMeter / 性能 / AWS の用語と動作", "設計レビュー"],
        ["05", "試験種別", "負荷・ストレス・スパイク・ソーク等", "計画の一文"],
        ["06", "全体観点", "業務・技術・観測・合否の軸", "観点漏れ防止"],
        ["07", "必要レポート", "残す成果物と HTML の使い方", "報告テンプレ"],
        ["08", "実行注意点", "結果を壊す落とし穴", "実施前読み合わせ"],
        ["09", "打ち込み位置", "どこから打つかのラン構成", "環境準備"],
        ["10", "API Gateway", "門の上限・429・29秒", "前段の扱い"],
        ["11", "ALB", "TargetResponseTime と 502/503/504", "標準エントリ"],
        ["12", "ECSタスク単位", "1タスク飽和と直線性", "サイジング"],
        ["13", "AuroraSv2", "ACU・バッファ・ポーズ", "DB 伸縮"],
        ["14", "RDS Proxy", "多重化とピン留め", "接続設計"],
        ["15", "ElastiCache", "ヒット・ホットキー・接続", "キャッシュ"],
        ["16", "帳票サーバ", "同期・並列・タイムアウト入れ子", "印刷室"],
        ["17", "切り分け手順", "遅さの分解レシピ", "試験中の判断"],
        ["18", "監視メトリクス", "段ごとの CloudWatch", "ダッシュボード"],
        ["19", "合否と失敗", "判定例（合格/条件付き/不合格）", "報告"],
        ["19b", "失敗の見分け", "症状→疑う段→誤薬→正攻法", "試験中の切り分け"],
        ["20", "チェックリスト", "計画・実行・事後の箱", "当日"],
        ["21", "たとえ話集", "小学生向け説明の素材", "ステークホルダ説明"],
        ["22", "コマンド集", "JMeter と確認クエリ", "オペ"],
        ["23", "参考情報", "一次情報 URL と確認事項", "数値の根拠"],
    ]
    put_rows(ws, 5, rows)


def principles(wb: Workbook) -> None:
    ws = sheet_base(wb, "02_原則", GOLD)
    apply_widths(ws, {1: 8, 2: 28, 3: 52, 4: 52, 5: 28})
    banner(ws, 1, 5, "全体性能試験の原則（これだけ外さない）", NAVY)
    note(ws, 2, 5, "原則に反するランは、数字が出ても「別のものを測っている」。報告書の「測っていない範囲」に必ず書く。")
    headers(ws, 4, ["#", "原則", "意味", "破ると起きること", "確認"])
    put_rows(
        ws,
        5,
        [
            [1, "測りたい段の直前から打つ", "前段は前段の試験", "API Gateway の 10,000 RPS や 429 をアプリ能力と誤認", "ラン表"],
            [2, "ECS はタスク単位が先", "1 タスク飽和 → 倍数で直線性", "平均 CPU 40% なのに一部タスクが火を噴いている", "Desired=1 ラン"],
            [3, "種類を混ぜない", "負荷・スパイク・ソークは別ラン", "スケール未起動の 5xx でアプリ不合格", "計画の一文"],
            [4, "JMeter は非 GUI", "測る機械が詰まると全部遅い", "サーバ CPU 暇なのにタイムアウト", "ジェネレータ CPU<70%"],
            [5, "平均よりパーセンタイル", "p95/p99 がユーザ体感", "平均 200ms・p99 8s を合格にする", "Statistics"],
            [6, "ウォームアップ除外", "JIT・接続・キャッシュ・ACU を温める", "朝一の遅さを定常 SLA にする／逆に温めすぎ", "除外分を明記"],
            [7, "サーバ側時計を同時に見る", "JMeter は外側の秒数しか知らない", "ECS を増やしても直らない共有天井を見逃す", "18_監視"],
            [8, "合否は事前の数値", "事後の感想で合格にしない", "再試験のたびに基準が動く", "19_合否"],
            [9, "帳票は別天井", "印刷室は ECS を増やしても増えない", "全体 TPS 不足の原因をアプリに誤る", "あり/なし分離"],
            [10, "伸縮はペナルティを分離", "ACU・ALB・ECS 起動の遅れは別指標", "スパイク不合格をチューニング対象と誤る", "固定ラン"],
            [11, "データと認証を本番相当に", "同じキー・1 ユーザ・期限切れ JWT", "ヒット率 100% や後半 401", "CSV 設計"],
            [12, "一次資料を残す", "JTL・ログ・UTC・プロパティ", "再現不能で改善検証ができない", "07_レポート"],
        ],
    )


def history(wb: Workbook) -> None:
    ws = sheet_base(wb, "03_歴史年表", SLATE)
    apply_widths(ws, {1: 16, 2: 28, 3: 40, 4: 50, 5: 36})
    banner(ws, 1, 5, "登場の歴史と、必要になった経緯", NAVY)
    note(ws, 2, 5, "遅さはいつも新しい共有層に移る。層を足したら、その層の試験観点を足す。")
    headers(ws, 4, ["時期", "何が起きたか", "なぜ必要になったか", "性能の新しい問題", "今の試験への残響"])
    put_rows(
        ws,
        5,
        [
            ["1960-80s", "銀行・航空のオンライン", "同時に大勢が同じ計算機", "同時接続・応答時間・ピーク", "TPC 的なトランザクション思考"],
            ["1990s Web", "ブラウザが世界から叩く", "Slashdot 効果で昨日まで静か・今日壊滅", "スケールとキャッシュ", "負荷ツールの一般化"],
            ["1998", "JMeter 原型（Mazzocchi）", "Java サーバを Java で測る", "内製ベンチ", "GUI シナリオ文化の起点"],
            ["2001", "Apache Jakarta JMeter", "Web 負荷の標準ツール化", "録画と再生、相関漏れ", "JMX 資産が残る理由"],
            ["2003-11", "memcached / Redis / ElastiCache", "同じ読込で DB が死ぬ", "ヒット率・ホットキー", "ホット/コールド別ラン"],
            ["2011", "JMeter が Apache TLP", "企業利用の定着", "分散試験", "既存資産で JMeter が残る"],
            ["2014-17", "ECS / Fargate", "箱の単位でデプロイと課金", "タスクサイズと台数", "1 タスク能力がサイジングの核"],
            ["2015", "API Gateway", "API の門をマネージドに", "アカウント RPS・429・29秒", "門と厨房を分けて測る"],
            ["2016", "ALB / JMeter 3.0 HTML", "コンテナへの L7 振り分け、報告の標準化", "502/503/504 の意味分化", "TargetResponseTime が外測の主"],
            ["2018-22", "Aurora Serverless v1→v2", "ピーク合わせ常時大型が無駄", "ACU 伸縮・バッファ・ポーズ", "固定ランと伸縮ラン"],
            ["2019", "RDS Proxy", "コンテナ/Lambda の接続嵐", "ピン留め・三重プール", "Pinned メトリクス必須"],
            ["2024-12", "ECS Container Insights enhanced", "タスク/コンテナまで自動観測", "平均が隠す偏りを可視化", "タスク単位メトリクス"],
            ["2024-26", "ALB keepalive 寿命、Aurora 高速スケール、ECS 20秒メトリクス", "突発負荷とエージェントワークロード", "追いつくまでの空白は残る", "ランプとスパイクを混ぜない"],
            ["2026 現在", "JMeter 5.6.3 が公式最新安定", "k6 等も増えるが JMX 資産は現役", "Coordinated Omission、非 GUI", "CLI + HTML + CloudWatch"],
        ],
    )


def glossary(wb: Workbook) -> None:
    ws = sheet_base(wb, "04_用語辞典", TEAL)
    apply_widths(ws, {1: 22, 2: 16, 3: 36, 4: 50, 5: 40, 6: 22})
    banner(ws, 1, 6, "用語辞典（名前 → 小学生向け → 動作原理 → 試験での使い方）", NAVY)
    note(ws, 2, 6, "JMeter 用語と性能一般、AWS 層の用語を混在させている。カテゴリでフィルタすること。")
    headers(ws, 4, ["用語", "カテゴリ", "小学生向け", "動作原理", "試験での使い方", "よくある誤解"])
    rows = [
        ["Test Plan / JMX", "JMeter", "試験の設計図", "XML でツリーを保存。機能モードは本文保持でメモリ死", "版管理する。機能モード OFF", "JMX があれば再現できる（プロパティ欠ける）"],
        ["Thread Group", "JMeter", "お客役の集団", "スレッド数・ランプ・ループ/時間", "同時ユーザか到着率かを先に決める", "スレッド数＝TPS"],
        ["Open Model Thread Group", "JMeter", "1秒に何人来るかで決める", "クローズド（人数固定）だと遅いとき TPS が自然減", "本番がオープンならこちら", "いつも Thread Group で足りる"],
        ["Ramp-up", "JMeter", "少しずつ増やす", "100人/100秒なら約1秒に1人", "オートスケールを間に合わせる", "短いほど「厳しい」良い試験"],
        ["Sampler", "JMeter", "注文を出す動作", "HTTP 等を送り SampleResult を作る", "ラベルを業務名で安定させる", "ブラウザと同じ"],
        ["Connect Time", "JMeter", "電話をかけるまで", "TCP/TLS 握手", "毎回長い＝Keep-Alive 失敗か枯渇", "アプリの遅さ"],
        ["Latency", "JMeter", "相手が「はい」と言うまで", "最初のバイトまで（TTFB に近い）", "PDF は Latency 短・Elapsed 長が正常", "Elapsed と同じ"],
        ["Elapsed / Response Time", "JMeter", "料理が全部届くまで", "最後のバイトまで", "SLA の主指標（パーセンタイル）", "サーバ CPU 時間"],
        ["Transaction Controller", "JMeter", "一連の注文を 1 業務にする", "親サンプルを TPS に出せる", "タイマーを含むかで SLA が変わる", "Hits/s と同じ"],
        ["Think time / Timer", "JMeter", "メニューを見る時間", "サンプラー前に待つ", "ゼロはロボット攻撃", "Pacing と同じ"],
        ["CSV Data Set", "JMeter", "名簿から名前を読む", "スレッドごとに行を割り当て", "Recycle と Stop を意図的に", "1 ユーザ使い回しで十分"],
        ["Cookie Manager", "JMeter", "整理券を持って歩く", "スレッドごとに Cookie 保持", "ほぼ必須", "Header にベタ書きで足りる"],
        ["Correlation / Extractor", "JMeter", "受け取った番号を次に使う", "JSON/正規表現/境界で抽出", "デバッグ負荷で 100% 成功させてから本試験", "録画のまま再生"],
        ["Assertion", "JMeter", "届いた料理が注文どおりか", "コード・本文・時間で success を上書き", "本試験は軽く。業務 200 エラーを拾う", "HTTP 200＝成功"],
        ["Listener", "JMeter", "結果のテレビ", "実行中 GUI リスナーは CPU を食う", "本試験は -l のみ。Tree 禁止", "Aggregate を付けて回す"],
        ["JTL", "JMeter", "全部のストップウォッチ帳", "CSV 推奨。全レポートの原料", "保管・マスキング", "HTML があれば捨ててよい"],
        ["HTML Dashboard", "JMeter", "説明用の絵", "3.0 以降の標準。APDEX と時系列", "一次資料ではない", "これだけで切り分けできる"],
        ["APDEX", "JMeter", "満足した人の割合", "satisfied / tolerated 閾値で 0〜1", "SLA に合わせて閾値変更（既定 0.5/1.5s は危険）", "既定のまま帳票を評価"],
        ["Throughput", "性能", "1秒に何件終わったか", "件数/秒。バイト/秒は別", "目標 TPS を先に書く", "帯域と同じ"],
        ["Percentile p95", "性能", "100人中95人目の待ち", "裾野の遅さ", "SLA は p95 か p99", "平均で代用"],
        ["Coordinated Omission", "性能", "遅い時間に注文が減って平均が良く見える", "クローズドワークロードの性質", "到着率モデルとパーセンタイル", "JMeter のバグ"],
        ["Warm-up", "性能", "準備運動", "JIT・プール・キャッシュ・ACU", "統計から除外", "短いほど効率的"],
        ["Saturation", "性能", "これ以上人を増やしても皿が増えない", "TPS 頭打ち＋遅延上昇", "1 タスク能力の定義点", "CPU 100% だけが飽和"],
        ["SLA / SLO / SLI", "性能", "約束 / 目標 / ものさし", "外向け・内向け・計測値", "合否表に写す", "全部同じ"],
        ["Workload mix", "性能", "メニューの割合", "参照70/更新20/帳票10 など", "ピーク1時間のログから", "全 API 均等"],
        ["Keep-Alive", "HTTP", "電話を切らずに次の話", "接続再利用で TLS を省略", "本番プロトコルに合わせる", "切った方が公平"],
        ["API Gateway", "AWS", "校門の警備", "認可・スロットル・統合。既定 10,000 RPS", "アプリ試験では外す", "無限にスケール"],
        ["429", "AWS", "門が満員", "トークンバケツが空", "アプリバグではない", "5xx と同じ扱い"],
        ["Integration timeout 29s", "AWS", "門は29秒で追い出す", "REST 既定。HTTP API は30秒", "帳票同期の壁", "アプリ設定で延びる"],
        ["ALB", "AWS", "配膳口", "L7 振り分け。idle 既定60s", "アプリ試験の標準エントリ", "Classic ELB と同じウォーム必須"],
        ["TargetResponseTime", "AWS", "窓の奥の調理時間", "target_processing_time", "アプリ遅延のきれいな外測", "JMeter Elapsed と同じ"],
        ["ELB 502/503/504", "AWS", "渡し方の失敗の種類", "切断/健全ゼロかALB不足/待ち過ぎ", "コードで原因段が違う", "全部アプリ障害"],
        ["ECS Task", "AWS", "調理員ひとり", "CPU/メモリの課金と制限の単位", "1 タスク飽和が先", "サービス＝1 プロセス"],
        ["CPUUtilization (ECS)", "AWS", "予約した力のうち使った割合", "サービス平均は偏りを隠す", "タスク次元で見る", "40% なら余裕"],
        ["RequestCountPerTarget", "AWS", "一人あたり何皿", "ALB がターゲットごとに数える", "水平スケールの密度", "クラスタ全体の RPS"],
        ["Fargate", "AWS", "厨房の建物を借りない", "タスクサイズの組が固定", "垂直スケールは組の変更", "EC2 と同じ無制限 CPU"],
        ["ACU", "AWS", "倉庫の床の広さ", "約 2GiB メモリ＋CPU＋ネット", "min/max と実測 ServerlessDatabaseCapacity", "vCPU と同じ"],
        ["Aurora auto-pause", "AWS", "倉庫を畳んで電気を消す", "min=0。復帰に時間がかかる", "本試験中は禁止", "v2 は常にポーズしない"],
        ["RDS Proxy multiplexing", "AWS", "席の使い回し", "トランザクション単位で DB 接続を再利用", "ピン留め率で効果を判定", "入れれば接続無限"],
        ["Pinning", "AWS", "専用ロッカーを使い席が固定", "SET・一時表・SQL>16KB 等", "Pinned メトリクス必須", "ログに必ず出るエラー"],
        ["EngineCPUUtilization", "AWS", "棚番一人の忙しさ", "Redis/Valkey コマンド処理の主スレッド", "ノード CPU より先に見る", "CPUUtilization と同じ"],
        ["Cache stampede", "AWS", "棚が空で全員が倉庫へ", "一斉ミスで DB へ雪崩", "コールドランとロック戦略", "ヒット率が高いから起きない"],
        ["帳票同期", "業務", "注文の最後に成績表を刷って渡す", "スレッドとタイムアウトを占有", "あり/なし分離、29秒監視", "ECS を増やせば速くなる"],
        ["JSR223 Groovy", "JMeter", "賢い計算メモ", "Compilable。キャッシュ時は vars.get", "BeanShell は負荷向きでない", "${var} 埋め込みでキャッシュ"],
        ["Distributed test", "JMeter", "お客役マシンを複数", "Master/Worker または独立実行＋JTL 結合", "時計同期と帯域", "1 台で無限"],
        ["Backend Listener", "JMeter", "生中継カメラ", "InfluxDB 等へリアルタイム", "長時間試験の監視", "GUI リスナーの代わりに Tree"],
    ]
    put_rows(ws, 5, rows)
    ws.auto_filter.ref = f"A4:F{4 + len(rows)}"


def test_types(wb: Workbook) -> None:
    ws = sheet_base(wb, "05_試験種別", ORANGE)
    apply_widths(ws, {1: 18, 2: 28, 3: 36, 4: 36, 5: 32, 6: 22})
    banner(ws, 1, 6, "試験の種類（今日のランはどれか、を一文で書く）", NAVY)
    note(ws, 2, 6, "全体パフォーマンステストの中心は負荷試験。スパイクとソークは必要なら別ラン。")
    headers(ws, 4, ["種類", "小学生向け", "証明すること", "標準の形", "混ぜると壊れるもの", "成果物"])
    put_rows(
        ws,
        5,
        [
            ["スモーク性能", "火が付くか", "スクリプトと観測が生きている", "1〜5 VU 数分", "本試験の数字として使う", "成功ログ"],
            ["負荷 (Load)", "いつもの給食時間", "目標負荷で SLA", "ランプ＋定常 20〜60分", "短いランプでスケール未起動", "合否表"],
            ["1タスク飽和", "選手一人のタイム", "単位能力", "Desired=1、スケール停止", "オートスケールが介入", "TPS/CPU 曲線"],
            ["スケールアウト", "人数を倍にしたら皿も倍か", "直線性", "2/4/8 タスク", "共有天井を無視して ECS 増設", "直線性グラフ"],
            ["ストレス", "限界まで増やす", "壊れ方と回復", "段階的に上限超え", "データ破壊の計画なし", "崩壊点"],
            ["スパイク", "突然どっと来る", "バースト耐性", "短時間 2〜5 倍", "API GW バーストや ALB 503 をアプリ不合格", "429/503 時刻"],
            ["ソーク/耐久", "何時間も出し続ける", "リーク・GC・ディスク", "数時間〜一晩", "短時間平均で合格", "右肩上がり有無"],
            ["キャパシティ", "何人まで安全か", "増設単位", "飽和点の外挿を実測で検証", "平均だけで外挿", "推奨タスク数"],
            ["フェイルオーバー", "壊したときの速さ", "切替とエラー率", "意図的に落とす", "通常ラン中の偶然障害", "RTO 実測"],
            ["E2E 体感", "校門から食べるまで", "ユーザ体感", "Gateway 通し 1 本", "原因特定に使う", "差分 ms"],
            ["コンポーネント単体", "印刷室だけ", "その部屋の天井", "直叩き", "全体ミックスと平均する", "部屋の SLA"],
        ],
    )


def viewpoints(wb: Workbook) -> None:
    ws = sheet_base(wb, "06_全体観点", TEAL)
    apply_widths(ws, {1: 16, 2: 28, 3: 44, 4: 44, 5: 28})
    banner(ws, 1, 5, "全体パフォーマンステストで必要な観点", NAVY)
    note(ws, 2, 5, "観点は測る軸。軸が無いと数字の山になる。各軸に「見るもの」「合格の形」「担当」を割り当てる。")
    headers(ws, 4, ["軸", "観点", "見ること", "欠けると", "主担当"])
    put_rows(
        ws,
        5,
        [
            ["目的", "リリース判定かサイジングか", "一文の目的", "打ち込み位置がぶれる", "PM / 性能"],
            ["業務", "ピークの形とミックス", "同時ユーザと TPS、帳票比率、データ量", "スリッパで運動会を語る", "業務 / アプリ"],
            ["クライアント", "プロトコルと認証", "HTTP/1.1 vs2、JWT 期限、Keep-Alive", "後半 401、TLS 嵐", "アプリ"],
            ["打ち込み", "どの段の試験か", "ALB 標準、GW 別、1 タスク別", "門の試験を厨房の成績にする", "性能"],
            ["ECS", "タスク単位", "1 タスク飽和、偏り、プール×台数", "平均 40% の安心", "アプリ / SRE"],
            ["キャッシュ", "ヒットとホットキー", "ホット/コールド、EngineCPU", "楽観ヒットや stampede 見逃し", "アプリ"],
            ["接続", "Proxy と max_connections", "ピン留め、三重プール", "スケールアウトで DB 死亡", "DBA"],
            ["DB 伸縮", "ACU 固定 vs 伸縮", "バッファ、上限張り付き", "スケール待ちをバグと呼ぶ", "DBA"],
            ["帳票", "同期と並列", "キュー、29秒、スレッド飢餓", "ECS 増設という誤薬", "アプリ"],
            ["観測", "外側と内側", "JMeter＋CloudWatch＋トレース", "外の秒数しか残らない", "SRE"],
            ["データ", "偏りと倫理", "ユニーク、マスキング、本番量", "制約違反や個人情報", "データ"],
            ["合否", "事前数値", "p95、エラー定義、除外区間", "感想で合格", "全体"],
            ["安全", "課金と破壊", "スタブ、上限、本番禁止", "請求事故・迷惑メール数千", "PM"],
            ["再現", "版とコマンド", "JMX、J 値、UTC、データ版", "直ったかが分からない", "性能"],
        ],
    )


def reports(wb: Workbook) -> None:
    ws = sheet_base(wb, "07_必要レポート", GREEN)
    apply_widths(ws, {1: 8, 2: 28, 3: 40, 4: 40, 5: 22, 6: 16})
    banner(ws, 1, 6, "必要なレポートと成果物", NAVY)
    note(ws, 2, 6, "HTML ダッシュボードは説明用。切り分けと再現は JTL と CloudWatch。結論は 1 ページ目。")
    headers(ws, 4, ["必須", "成果物", "中身", "使い方", "保管", "一次/二次"])
    put_rows(
        ws,
        5,
        [
            ["必須", "テスト計画書", "目的・範囲・打ち込み・負荷・合否", "キックオフ承認", "リポジトリ", "一次"],
            ["必須", "JMX と版", "実行シナリオ", "再現", "Git タグ", "一次"],
            ["必須", "user.properties / -J", "スレッド・ホスト・時間", "再現", "実行ログ横", "一次"],
            ["必須", "JTL (CSV)", "全サンプル", "再集計の唯一の原料", "長期", "一次"],
            ["必須", "jmeter.log", "クライアント例外", "ジェネレータ死因", "長期", "一次"],
            ["必須", "HTML Dashboard", "APDEX, Statistics, Over Time, Errors", "関係者説明。貼る最小は負荷・TPS・p95・エラー・Connect", "長期", "二次"],
            ["必須", "開始終了 UTC", "時計合わせ", "CloudWatch 接合", "報告書", "一次"],
            ["必須", "ジェネレータ CPU/MEM/NET", "クライアント健全", "70% 超は無効疑い", "試験フォルダ", "一次"],
            ["必須", "CloudWatch 画面/CSV", "各段の時系列", "ボトルネック位置", "試験フォルダ", "一次"],
            ["必須", "合否レポート 1 ページ", "合格/条件付き/不合格", "リリース判断", "チケット", "二次"],
            ["必須", "既知問題とアクション", "誰がいつ直す", "改善に変える", "チケット", "二次"],
            ["推奨", "トレース抜粋", "p99 の内訳 span", "コード行まで", "一時", "一次"],
            ["推奨", "ALB アクセスログ抜粋", "三段時計とターゲット IP", "偏りと 504", "一時", "一次"],
            ["推奨", "1 タスク曲線", "TPS vs CPU vs p95", "サイジング根拠", "計画改訂", "二次"],
            ["推奨", "直線性表", "1/2/4/8 タスク", "共有天井の証明", "計画改訂", "二次"],
            ["条件", "スパイク時刻表", "429/503 とスケール開始", "バースト耐性", "別ラン", "一次"],
            ["条件", "ソーク傾き", "ヒープ・接続の右肩", "リーク", "別ラン", "一次"],
            ["条件", "帳票キュー", "in-flight と待ち", "印刷室天井", "別ラン", "一次"],
        ],
        filter_header=4,
    )
    r = section_title(ws, 25, 6, "HTML Dashboard の見方（公式が出すもの）", GOLD)
    headers_row = r
    for i, t in enumerate(["グラフ/表", "含むもの", "見るポイント", "注意", "貼るか", "-"], 1):
        cell = ws.cell(headers_row, i, t)
        cell.fill = fill(TEAL)
        cell.font = font(True, WHITE, 11)
        cell.alignment = CENTER
        cell.border = THIN
    put_rows(
        ws,
        headers_row + 1,
        [
            ["APDEX", "トランザクション別 0〜1", "閾値を SLA に変更", "既定 500/1500ms は帳票を全不満にする", "条件付き", ""],
            ["Requests Summary", "成功/失敗円", "失敗が無視できない割合か", "トランザクション親を含まない", "はい", ""],
            ["Statistics", "平均・分位・TPS・エラー", "p95/p99 とエラー率", "平均だけで語らない", "はい", ""],
            ["Errors / Top5", "コードとメッセージ", "401 と 429 と 504 を分けて数える", "業務 200 エラーはアサーション無しだと消える", "はい", ""],
            ["Active Threads Over Time", "負荷プロファイル", "ランプと定常が意図どおりか", "落ちていたらジェネレータ死", "はい", ""],
            ["Response times Over Time", "時系列遅延", "階段状ならスケール待ち", "粒度 <1s は公式非推奨", "はい", ""],
            ["Transactions per second", "業務 TPS", "目標維持", "Hits/s と別物", "はい", ""],
            ["Hits / Codes per second", "ヒットと HTTP コード", "429/5xx の時刻", "埋め込み資源の扱い", "条件", ""],
            ["Connect Time Over Time", "握手", "TLS 嵐・枯渇", "アプリ遅延と分離", "はい", ""],
            ["Latency vs Request", "負荷に対する待ち", "飽和の兆候", "オープン/クローズドで形が違う", "推奨", ""],
            ["Response time vs Threads", "人数対遅延", "スケール特性", "思考時間の影響", "推奨", ""],
        ],
        apply_filter=False,
    )


def cautions(wb: Workbook) -> None:
    ws = sheet_base(wb, "08_実行注意点", RED)
    apply_widths(ws, {1: 8, 2: 24, 3: 40, 4: 36, 5: 36, 6: 20})
    banner(ws, 1, 6, "実行する際の注意点（結果を壊すもの）", NAVY)
    note(ws, 2, 6, "公式ベストプラクティス: 最新版、正しいスレッド数、CLI、リスナー削減、CSV、アサーション最小、Groovy。")
    headers(ws, 4, ["重大度", "落とし穴", "兆候", "なぜ数字が嘘になるか", "対策", "確認先"])
    put_rows(
        ws,
        5,
        [
            ["致命", "GUI で本試験", "ジェネレータ CPU 高、描画遅", "測る機械がボトルネック", "必ず -n", "JMeter 公式"],
            ["致命", "View Results Tree を負荷中オン", "ヒープ急増", "本文保持", "デバッグ専用", "公式 16.7"],
            ["致命", "API Gateway 通しだけ", "429、課金、29s 504", "門の試験", "ALB 標準ラン", "10_API Gateway"],
            ["致命", "1 タスクを取らない", "平均 CPU だけ低い", "偏りと単位能力が不明", "Desired=1", "12_ECS"],
            ["高", "ランプ 1 秒でピーク", "起動中 503、ACU 階段", "伸縮試験が混入", "5〜15分ランプ＋スパイク別", "09_打ち込み"],
            ["高", "思考時間ゼロ", "非現実 TPS", "DB/帳票が過大破壊", "ログ由来の分布", "計画"],
            ["高", "データ 1 キー/1 ユーザ", "ロック or ヒット 100%", "楽観または悲観に偏る", "CSV 多様", "06_観点"],
            ["高", "JWT 期限切れ", "後半だけ 401", "負荷耐性ではなく認証寿命", "setUp で更新", "スクリプト"],
            ["高", "ジェネレータ過負荷", "Connect 増、サーバ暇", "無効試験", "CPU<70%、分散", "BlazeMeter 目安等"],
            ["高", "タイムアウト入れ子無視", "クライアントエラーなのにサーバ成功", "帳票が裏で完走", "層ごとの秒数表", "16_帳票"],
            ["高", "平均だけで合格", "p99 爆発", "1% の激怒", "p95/p99", "19_合否"],
            ["中", "機能モード ON", "OOM", "本文全保持", "OFF", "Test Plan"],
            ["中", "XML JTL", "ディスクと後処理が重い", "分析不能", "CSV、必要列だけ", "user.properties"],
            ["中", "BeanShell / ${var} 埋め込み Groovy", "CPU 高", "コンパイル不能・キャッシュ汚染", "JSR223 cache + vars.get", "公式 16.12"],
            ["中", "WAF が攻撃判定", "403 山", "セキュリティ試験が混入", "許可リストと別ラン", "セキュリティ"],
            ["中", "NTP ずれ", "グラフが噛み合わない", "切り分け不能", "UTC 統一", "全ホスト"],
            ["中", "隣の API とクォータ共有", "理由なき 429", "アカウント 10k RPS 共有", "試験枠の調整", "Service Quotas"],
            ["中", "キャッシュ温めすぎ", "DB 暇", "本番より楽観", "ヒット率目標を設定", "15_ElastiCache"],
            ["中", "ACU ポーズ", "初回だけ十数秒", "コールドを SLA に", "min>0、試験前ウォーム", "13_Aurora"],
            ["中", "アプリ keep-alive < ALB idle", "502", "切れた接続に ALB が送る", "アプリを長く", "11_ALB"],
            ["中", "帳票デバッグログ ON", "印刷だけ遅い", "測定を自分で遅くする", "OFF", "SVF KB"],
            ["中", "外部メール/SaaS 本番", "請求とレート制限", "隣を破壊", "スタブ＋遅延注入", "計画"],
            ["低", "jmeter.properties 直接編集", "次版移行で上書き", "設定喪失", "user.properties", "公式 16.14"],
            ["低", "HTML 粒度 <1s", "TPS グラフ崩壊", "公式が警告", ">=1000ms", "Dashboard 文書"],
        ],
    )


def injection(wb: Workbook) -> None:
    ws = sheet_base(wb, "09_打ち込み位置", ORANGE)
    apply_widths(ws, {1: 10, 2: 28, 3: 44, 4: 36, 5: 28, 6: 22})
    banner(ws, 1, 6, "打ち込み位置：測りたい段の直前から打つ", NAVY)
    note(ws, 2, 6, "標準の全体性能はランB（ALB）＋ランC（1タスク）。A は体感差分、F は門の試験。")
    headers(ws, 4, ["ラン", "打ち込み", "目的", "含むもの / 含まないもの", "合格の意味", "必須か"])
    put_rows(
        ws,
        5,
        [
            ["A 体感E2E", "Internet→WAF→API GW→ALB→ECS", "ユーザ体感", "前段全部。原因特定には使わない", "体感 SLA の参考", "1本"],
            ["B アプリ本体", "内部ALB→ECS", "アプリ＋直下依存", "門を含まない", "リリース判定の主", "必須"],
            ["C 1タスク", "ALB、Desired=1、スケール停止", "単位能力", "水平スケール無し", "サイジングの核", "必須"],
            ["D-1 キャッシュ有", "B と同じ、ホット", "本番相当ヒット", "コールドを含まない", "定常 SLA", "必須"],
            ["D-2 キャッシュ無", "バイパスまたは空", "DB 真の負荷", "ヒットを含まない", "朝一・障害時", "推奨"],
            ["D-3 Proxy無", "アプリ→DB 直", "ホップとピンの寄与", "Proxy 無し", "導入是非", "推奨"],
            ["E 帳票隔離", "帳票 API または直叩き", "印刷室天井", "OLTP を混ぜない", "帳票 SLA", "帳票があるなら必須"],
            ["B' 帳票なしミックス", "ALB、帳票 0%", "API 本体", "印刷室を含まない", "ECS 増設が効く範囲", "必須"],
            ["F ゲートウェイ", "API GW 通し", "429・認可・29s", "厨房能力の主測定にしない", "門の設定", "別チーム/別日"],
            ["G スパイク", "B の形で急増", "バースト", "定常と混ぜない", "伸縮の空白", "任意"],
            ["H ソーク", "B を数時間", "リーク", "短時間と混ぜない", "耐久", "リリース前"],
            ["I ACU固定", "B、min=max ACU", "伸縮無しの DB 能力", "スケール待ち無し", "スキーマ/SQL の真", "推奨"],
        ],
    )
    r = section_title(ws, 19, 6, "ジェネレータの置き場所", GOLD)
    headers2 = r
    for i, t in enumerate(["置き場所", "測れるもの", "乗ってしまうもの", "向き", "注意", "-"], 1):
        c = ws.cell(headers2, i, t)
        c.fill = fill(TEAL)
        c.font = font(True, WHITE, 11)
        c.alignment = CENTER
        c.border = THIN
    put_rows(
        ws,
        headers2 + 1,
        [
            ["VPC 内（推奨・ランB）", "ALB 以降の能力", "インターネット往復は乗らない", "アプリ能力", "本番ユーザとの差を A で一度測る", "同一ホスト禁止"],
            ["パブリック Internet", "体感", "ISP、NAT、WAF、GW", "E2E", "許可リスト", "Aラン向き"],
            ["同じホストにアプリと JMeter", "何も正しくない", "CPU 奪い合い", "禁止", "別マシンへ移す", "無効試験"],
            ["別リージョン", "広域遅延", "リージョン差", "DR 以外は避ける", "同一リージョン別AZ", "体感差はAで測る"],
        ],
        apply_filter=False,
    )


def apigw(wb: Workbook) -> None:
    ws = sheet_base(wb, "10_API Gateway", NAVY)
    apply_widths(ws, {1: 22, 2: 50, 3: 50, 4: 28})
    banner(ws, 1, 4, "API Gateway：門の試験にしない", NAVY)
    note(ws, 2, 4, "たとえ: 校門の警備員。厨房が速くても門が狭いと外では「学校が遅い」。2015年、API の認証・キー・レート制限の自作が限界になりマネージドの門が標準になった。")
    r = section_title(ws, 4, 4, "動作原理", GOLD)
    headers(ws, r, ["要素", "動き", "試験への効き", "既定・上限（確認必須）"])
    # headers() sets freeze and autofilter; subsequent tables on same sheet will be fine enough
    put_rows(
        ws,
        r + 1,
        [
            ["受け口", "HTTPS を受けステージ/ルート決定", "エッジとリージョナルでレイテンシ差", "ペイロード 10MB"],
            ["オーソライザ", "IAM / Cognito / Lambda", "キャッシュ無しだと先に死ぬ", "結果キャッシュ TTL"],
            ["スロットル", "トークンバケツ。空なら 429", "アカウント×リージョンの全 API 共有", "10,000 RPS、バースト最大 5,000（一部リージョン 2,500/1,250）"],
            ["Usage Plan / API Key", "クライアント別のバケツ", "試験キーと本番キーを分ける", "プラン設定"],
            ["統合", "HTTP / Lambda / VPC Link→ALB", "IntegrationLatency が後ろの時間", "VPC Link v2 で private ALB 直結可"],
            ["時計", "Latency 全体、IntegrationLatency 後ろ", "差が大きいと門自身が重い", "CloudWatch"],
            ["時間切れ", "統合タイムアウト", "帳票同期が 504。アプリは成功していることあり", "REST 29s（Regional/private は申請で延長可）、HTTP API 30s は不可"],
            ["課金", "リクエスト課金", "高負荷試験の費用を支配", "別アカウント/ステージ"],
        ],
    )
    r2 = section_title(ws, 14, 4, "観点・やり方・やってはいけないこと", GOLD)
    for i, t in enumerate(["観点", "やること", "やってはいけない", "証拠"], 1):
        c = ws.cell(r2, i, t)
        c.fill = fill(TEAL)
        c.font = font(True, WHITE, 11)
        c.alignment = CENTER
        c.border = THIN
    put_rows(
        ws,
        r2 + 1,
        [
            ["アプリ能力", "ALB 直打ちを主にする", "GW 通しの TPS をアプリ能力と呼ぶ", "ランB"],
            ["門の上限", "Service Quotas を計画に写す", "10,000 ユーザと 10,000 RPS を混同", "クォータ画面"],
            ["429", "コードをアプリエラーから除外して数える", "リトライ嵐でさらにバケツを空にする", "アクセスログ status"],
            ["認可", "トークン更新とオーソライザキャッシュ", "1 トークンを 2 時間試験", "401 時系列"],
            ["29秒", "帳票 p99 と突合", "GW 経由で巨大 PDF 同期", "504 とアプリ成功ログ"],
            ["隣のシステム", "試験枠を調整", "同アカウント同時間に別負荷", "Count 合計"],
            ["マッピング", "大きな VTL の差を Integration と比較", "変換をアプリ遅延と呼ぶ", "Latency-IntegrationLatency"],
        ],
        apply_filter=False,
    )


def alb(wb: Workbook) -> None:
    ws = sheet_base(wb, "11_ALB", TEAL)
    apply_widths(ws, {1: 24, 2: 48, 3: 48, 4: 30})
    banner(ws, 1, 4, "ALB：アプリ試験の標準エントリ", NAVY)
    note(ws, 2, 4, "たとえ: 配膳口。2016年、コンテナとパスベースのために L7 特化で登場。Classic ELB の事前ウォーム文化は薄れたが、ゼロから数万 RPS を数秒では 503 があり得る。")
    r = section_title(ws, 4, 4, "時計とエラーコード", GOLD)
    headers(ws, r, ["名前", "意味", "遅い/出るとき", "対策"])
    put_rows(
        ws,
        r + 1,
        [
            ["request_processing_time", "受けてからタスクへ渡すまで", "ALB 自身。通常小さい", "異常なら ALB スケールやルール"],
            ["target_processing_time / TargetResponseTime", "タスクの仕事", "アプリ・依存の遅さ", "タスク単位とトレース"],
            ["response_processing_time", "クライアントへ書き戻し", "巨大ボディ、遅いクライアント", "帳票 PDF サイズ"],
            ["JMeter Elapsed との差", "インターネット＋TLS＋GW", "差が大きいなら経路", "ランAとBの差分"],
            ["idle timeout 既定 60s（1-4000）", "無通信で切る", "帳票同期 504。ログ値が idle と一致", "短縮するか非同期化"],
            ["HTTP client keepalive duration（2024〜、既定3600s）", "接続の最大寿命", "長時間ソークでの再接続", "ジェネレータも追従"],
            ["アプリ keep-alive > ALB idle", "切れた接続への送り防止", "502", "公式推奨"],
            ["502", "タスクが乱暴に切断", "プロセス死、keep-alive 不整合", "タスクログ、graceful"],
            ["503", "健全ターゲットゼロ or ALB 不足", "起動中、急ランプ", "ランプ、ヘルス、ウォーム"],
            ["504", "idle まで返らず", "アプリ/帳票/DB 待ち", "層の秒数表"],
            ["UnHealthyHostCount", "ヘルス失敗", "重いヘルスが雪崩を呼ぶ", "軽いヘルス"],
            ["RequestCountPerTarget", "1 タスクあたり皿数", "サイジング密度", "1 タスク飽和と一致させる"],
            ["スロースタート", "新規ターゲットに徐々に振る", "スケール直後 p99 悪化", "別指標"],
            ["sticky session", "同じ人を同じ窓へ", "偏り、平均 CPU 低・p99 高", "無効化して比較"],
        ],
    )


def ecs(wb: Workbook) -> None:
    ws = sheet_base(wb, "12_ECSタスク単位", GREEN)
    apply_widths(ws, {1: 22, 2: 48, 3: 48, 4: 30})
    banner(ws, 1, 4, "ECS：サービス平均ではなくタスク単位で取る", NAVY)
    note(
        ws,
        2,
        4,
        "たとえ: 調理員ひとり。平均「普通」でも一人だけ炎上することがある。2014 ECS、2017 Fargate で課金と制限の単位がタスクサイズになった。"
        "2024-12 Container Insights enhanced。2026-06 に 20 秒メトリクスでスケール開始が大幅に短縮（公式ベンチ）。",
    )
    r = section_title(ws, 4, 4, "なぜタスク単位か", GOLD)
    headers(ws, r, ["理由", "原理", "平均だけだと", "取り方"])
    put_rows(
        ws,
        r + 1,
        [
            ["サイジングの単位", "欲しい TPS / 1タスクTPS = 台数", "夜間 2 タスクに縮んだ朝に死ぬ", "Desired=1 で飽和点"],
            ["偏り", "sticky、長い帳票、GC、不均等", "平均 40% で安心", "TaskId 別 CPU"],
            ["垂直と水平", "0.25vCPU と 2vCPU では単一スレッド帳票が違う", "台数だけ議論", "サイズ変更ラン"],
            ["接続プール", "プール×タスク数 = DB 接続", "1 タスク 20 本を 50 タスクにすると 1000 本", "最大 Desired で計算"],
            ["障害の単位", "1 死 = 1/N のエラー", "N の根拠がない", "飽和点から N を決める"],
            ["オートスケール", "平均 70% 閾値は遅いタスクを隠す", "閾値未達で増えない", "RequestCountPerTarget も併用"],
        ],
    )
    r2 = section_title(ws, 12, 4, "推奨手順", GOLD)
    for i, t in enumerate(["手順", "操作", "記録するもの", "合格の見方"], 1):
        c = ws.cell(r2, i, t)
        c.fill = fill(TEAL)
        c.font = font(True, WHITE, 11)
        c.alignment = CENTER
        c.border = THIN
    put_rows(
        ws,
        r2 + 1,
        [
            ["1", "オートスケール停止、Desired=1", "開始 UTC、タスク ID、サイズ", "固定されている"],
            ["2", "ランプして TPS 上昇", "TPS、p95、CPU、メモリ、エラー", "曲線が描ける"],
            ["3", "飽和点（TPS 頭打ち＋p95 立ち上がり）", "その時の CPU/接続/帳票/ACU", "単位能力の定義"],
            ["4", "Desired=2,4,8。密度（RequestCountPerTarget）を揃える", "TPS 倍率、p95、共有メトリクス", "ほぼ直線なら水平可"],
            ["5", "直線でない", "Proxy、Aurora、キャッシュ、帳票", "共有天井。ECS 増は誤薬"],
            ["6", "スケールを戻しピーク到達時間", "起動、ヘルス、スロースタート、5xx", "スパイク要件と比較（別合格）"],
        ],
        apply_filter=False,
    )
    r3 = section_title(ws, 20, 4, "見るメトリクス", GOLD)
    for i, t in enumerate(["メトリクス", "次元", "読み方", "罠"], 1):
        c = ws.cell(r3, i, t)
        c.fill = fill(TEAL)
        c.font = font(True, WHITE, 11)
        c.alignment = CENTER
        c.border = THIN
    put_rows(
        ws,
        r3 + 1,
        [
            ["CPUUtilization / MemoryUtilization", "Service 平均", "全体の温度", "偏りを隠す。Linux は 100% 超も"],
            ["Running / Desired TaskCount", "Service", "起動遅れ", "Desired だけ上がり Running が追いつかない"],
            ["CpuUtilized per Task", "TaskId", "火を噴く一人", "カスタムメトリクス化は課金注意。Logs Insights 可"],
            ["RequestCountPerTarget", "ALB ターゲット", "一人あたり皿", "1 タスク飽和と単位を揃える"],
            ["JVM GC / heap / threads", "アプリ", "待ちの中身", "CPU 低・遅延高なら待ち"],
            ["ephemeral storage", "Fargate", "帳票一時ファイル", "ディスクフル 5xx"],
            ["高解像度 20s（2026〜）", "Service", "速いスケール用", "短いスパイクでも間に合わないことはある"],
        ],
        apply_filter=False,
    )


def aurora(wb: Workbook) -> None:
    ws = sheet_base(wb, "13_AuroraSv2", GOLD)
    apply_widths(ws, {1: 24, 2: 48, 3: 48, 4: 30})
    banner(ws, 1, 4, "Aurora Serverless v2：伸縮する倉庫", NAVY)
    note(
        ws,
        2,
        4,
        "v1（2018）は載せ替えで切れやすかった。v2（2022〜）は同じインスタンス内で ACU を細かく変える。"
        "ACU≒2GiB。min 0（ポーズ）〜 max 256。2026 年公式は 1 秒で +12 ACU などの高速化を発表。ゼロ遅延ではない。",
    )
    r = section_title(ws, 4, 4, "原理と試験観点", GOLD)
    headers(ws, r, ["項目", "原理", "試験でやること", "やってはいけない"])
    put_rows(
        ws,
        r + 1,
        [
            ["ACU", "メモリ・CPU・ネットの束。刻み 0.5 から", "ServerlessDatabaseCapacity を時系列で取る", "vCPU 枚数と同一視"],
            ["バッファプール", "よく使うデータはメモリ。min が小さいと捨てられる", "min をワークセットが載る値以上に", "min=0.5 のコールドを SLA に"],
            ["オートポーズ min=0", "暇なら畳む。復帰に時間がかかる（条件で十数秒）", "本試験中は min>0、事前ウォーム", "開発設定のまま本番相当試験"],
            ["固定ラン min=max", "プロビジョン相当", "アプリ＋スキーマの真の能力", "伸縮ランだけ"],
            ["伸縮ラン", "本番 min/max", "p95 の階段と ACU 上昇の遅れを別指標に", "スケール待ちを SQL バグと呼ぶ"],
            ["上限張り付き", "これ以上床が広がらない", "max 引き上げかプロビジョン/分割", "「サーバレスだから無限」"],
            ["単一ライターのロック", "ACU を増やしても直列は直列", "待機イベントを見る", "ACU 増だけを処方箋に"],
            ["アイドル接続", "スケールダウンしにくい", "Proxy とアプリプールの idle を短く", "接続張りっぱなし"],
            ["Reader", "読みを分散", "分析をライターに混ぜない", "読込試験をライターへ全部"],
            ["Performance Insights", "待機の内訳。最低 ACU 推奨あり（例: 2）", "現行ドキュメント確認", "0.5 ACU で PI 前提"],
        ],
    )


def rdsproxy(wb: Workbook) -> None:
    ws = sheet_base(wb, "14_RDS_Proxy", ORANGE)
    apply_widths(ws, {1: 24, 2: 48, 3: 48, 4: 30})
    banner(ws, 1, 4, "RDS Proxy：整理券機とピン留め", NAVY)
    note(
        ws,
        2,
        4,
        "2019 年頃、Lambda/コンテナの接続嵐とフェイルオーバー切断を和らげるために登場。"
        "トランザクション単位で DB 接続を使い回す。セッション状態が付くとピン留め（1:1）。SQL 16KB 超は全エンジンでピン（公式）。",
    )
    r = section_title(ws, 4, 4, "原理", GOLD)
    headers(ws, r, ["状態", "何が起きるか", "メトリクス", "試験アクション"])
    put_rows(
        ws,
        r + 1,
        [
            ["多重化（健全）", "外は多く、中は少ない。席の使い回し", "Client >> Database でも待ちが少ない", "あり/なし比較でレイテンシ差を見る"],
            ["ピン留め", "SET、一時表、巨大 SQL、一部 PREPARE 等で席が固定", "DatabaseConnectionsCurrentlySessionPinned", "本番相当の巨大 IN 句を再現"],
            ["三重プール", "アプリ池 + Proxy 池 + DB", "Borrowed とアプリ active", "idle を Proxy より短く。24h 寿命に注意"],
            ["ピン率が高い", "高い整理券機が延長コードになる", "Pinned ≈ Borrowed", "導入是非を再評価（外した事例あり）"],
            ["接続嵐", "タスク急増で DB プロセスが死ぬ", "ClientConnections 急増、DB conn 上限", "ここが Proxy の本領。Lambda 向き"],
            ["ヘルス SELECT 1", "アイドル判定を壊し接続が減らない", "conn が下がらない", "ヘルス間隔と idle の設計"],
            ["ヘッドルーム", "公式ダッシュボードはピークの 30% 余裕を推奨", "MaxDatabaseConnectionsAllowed", "試験で 100% 埋めない"],
        ],
    )
    r2 = section_title(ws, 13, 4, "ピン留めを誘発しやすいもの", GOLD)
    for i, t in enumerate(["原因", "なぜ", "見つけ方", "緩和"], 1):
        c = ws.cell(r2, i, t)
        c.fill = fill(TEAL)
        c.font = font(True, WHITE, 11)
        c.alignment = CENTER
        c.border = THIN
    put_rows(
        ws,
        r2 + 1,
        [
            ["SQL テキスト > 16KB", "公式・全エンジン", "ORM の巨大 IN、長い SELECT", "SQL 分割、ピン率監視"],
            ["SET / セッション変数", "接続ごとに状態が違うと再利用不能", "アプリ初期化、タイムゾーン", "Initialization query へ寄せる"],
            ["一時表", "そのセッション専用", "バッチ的 API", "アプリ側一時、またはピンを許容して台数計算"],
            ["PostgreSQL の変数", "MySQL よりピンになりやすい", "ドライバの session_track", "フィルタは MySQL 中心。PG は設計見直し"],
            ["DISCARD ALL リセット", "プール返却時", "一部ライブラリ既定", "リセット方法変更"],
        ],
        apply_filter=False,
    )


def elasticache(wb: Workbook) -> None:
    ws = sheet_base(wb, "15_ElastiCache", TEAL)
    apply_widths(ws, {1: 24, 2: 48, 3: 48, 4: 30})
    banner(ws, 1, 4, "ElastiCache：手元の棚（当たると世界が変わる）", NAVY)
    note(
        ws,
        2,
        4,
        "memcached 2003、Redis 2009、ElastiCache 2011、Valkey 選択肢 2024〜。接続確立（特に TLS）は GET より桁で高い。永続接続とプールが前提。",
    )
    r = section_title(ws, 4, 4, "観点", GOLD)
    headers(ws, r, ["観点", "原理", "ラン", "見るもの"])
    put_rows(
        ws,
        r + 1,
        [
            ["ホット", "本番相当ヒット", "ウォーム後定常", "HitRate、p95"],
            ["コールド", "全員が倉庫へ", "空または別キー", "DB CPU、stampede"],
            ["ホットキー", "人気の一瓶に行列", "意図的に同じ ID", "EngineCPU、単一シャード"],
            ["接続", "握手はコマンドより高い", "プールサイズ×タスク", "CurrConnections、NewConnections"],
            ["Cluster mode", "16384 スロット。クライアントは全シャードへ接続", "本番と同じモード", "MOVED、シャード偏り"],
            ["Eviction", "メモリ満杯で捨てる", "ソーク", "Evictions 後のミス"],
            ["TTL 一斉切れ", "同じ時刻に棚が空", "TTL を試験用に揃えすぎない", "ミスの波"],
            ["障害フォールバック", "棚が無いと DB", "ノード停止ラン", "サーキット、stale、単一フライト"],
            ["書込（セッション/ロック）", "レプリカ遅延と EngineCPU", "更新ミックス", "ReplicationLag"],
            ["JMeter 直 Redis", "アプリのプール挙動と別物", "原則アプリ経由", "直叩きは単体ベンチ"],
        ],
    )


def report_server(wb: Workbook) -> None:
    ws = sheet_base(wb, "16_帳票サーバ", RED)
    apply_widths(ws, {1: 24, 2: 48, 3: 48, 4: 30})
    banner(ws, 1, 4, "帳票サーバ：印刷室は ECS を増やしても増えない", NAVY)
    note(
        ws,
        2,
        4,
        "日本の業務では見積・請求 PDF が同期 HTTP で残りやすい。SVF はジョブ並列・プリンタスレッド、Jasper はメモリと virtualizer。"
        "同期は API Gateway 29s / ALB idle / アプリ・JMeter タイムアウトの入れ子に当たる。",
    )
    r = section_title(ws, 4, 4, "観点", GOLD)
    headers(ws, r, ["観点", "原理", "やり方", "誤薬"])
    put_rows(
        ws,
        r + 1,
        [
            ["あり/なし分離", "印刷室満杯で全体 TPS が頭打ち", "帳票 0% と 本番比率を別ラン", "ECS 増設"],
            ["隔離天井", "同時 N 台まで", "帳票だけ直叩きまたは専用 API", "OLTP と平均する"],
            ["サイズ種類", "1頁と 1000 行明細は別世界", "代表 3〜5 種", "1 頁 PDF だけで合格"],
            ["同期 SLA", "スレッド占有＋タイムアウト入れ子", "p99 が 29s/idle に余裕あるか", "GW 経由巨大バイナリ"],
            ["スレッド飢餓", "帳票待ちで他 API の糸が無くなる", "帳票オン時の参照系 p95", "同じタスクに詰める"],
            ["DB 二重打ち", "帳票側が再検索", "遅いクエリと接続奪い", "オンライン用インデックスだけ見る"],
            ["一時ディスク", "一時ファイル経年で劣化（SVF KB）", "ソークで件数/inode", "CPU だけ見る"],
            ["JVM", "ヒープと GC、フォント、画像", "ヒープダンプ方針", "デバッグログ ON のまま測定"],
            ["非同期化", "ジョブ ID＋S3 でオンラインから外す", "後のオンラインだけ再測定", "同期のまま台数増"],
            ["外部 SaaS", "相手のレート制限試験になる", "スタブ＋本番相当遅延", "本番 SaaS を殴る"],
            ["製品並列上限", "UCX 並列、PrinterThread 等", "設定値とキュー待ち", "無限並列だと思う"],
        ],
    )
    r2 = section_title(ws, 17, 4, "タイムアウト入れ子（外側が先に切る）", GOLD)
    for i, t in enumerate(["層", "既定の目安", "切れた見え方", "帳票への意味"], 1):
        c = ws.cell(r2, i, t)
        c.fill = fill(TEAL)
        c.font = font(True, WHITE, 11)
        c.alignment = CENTER
        c.border = THIN
    put_rows(
        ws,
        r2 + 1,
        [
            ["JMeter", "設定次第", "クライアント失敗。サーバは成功のことも", "アサーションとサーバログを突合"],
            ["API Gateway REST", "29 秒（延長は Regional/private のみ）", "504", "同期帳票のハード上限になりやすい"],
            ["API Gateway HTTP API", "30 秒（上げられない）", "504", "長い帳票は設計不適合"],
            ["ALB idle", "60 秒（1〜4000）", "504、target_processing_time≒idle", "進捗バイトを送るか非同期"],
            ["アプリ（Tomcat 等）", "設定", "500 / 切断 → ALB 502", "帳票専用タイムアウト"],
            ["帳票ジョブ", "製品設定", "キュー待ちが Elapsed に乗る", "並列上限の証拠"],
        ],
        apply_filter=False,
    )


def correlation(wb: Workbook) -> None:
    ws = sheet_base(wb, "17_切り分け手順", SLATE)
    apply_widths(ws, {1: 8, 2: 36, 3: 44, 4: 44, 5: 28})
    banner(ws, 1, 5, "遅さを分解する（遠足の遅れをバスと徒歩に分ける）", NAVY)
    note(
        ws,
        2,
        5,
        "Elapsed ≒ 往復 + (GW Latency-IntegrationLatency) + 認可 + ALB request + TargetResponseTime(アプリ+キャッシュ+Proxy+DB+帳票) + response。",
    )
    headers(ws, 4, ["順", "操作", "差が出たら原因", "差が無ければ", "次"])
    put_rows(
        ws,
        5,
        [
            ["1", "ALB 直と GW 通しを同じ負荷で", "門（スロットル、認可、マッピング）", "門は主犯ではない", "2"],
            ["2", "キャッシュ有無", "ヒット率・ホットキー・stampede", "キャッシュは主犯ではない", "3"],
            ["3", "帳票あり/なし", "印刷室とスレッド飢餓", "帳票は主犯ではない", "4"],
            ["4", "Desired=1 で CPU を見る", "CPU 高＝アプリ計算。CPU 低・遅い＝待ち", "待ちなら 5", "5"],
            ["5", "DB 待機イベント", "CPU＝ACU、Lock＝SQL、IO＝バッファ", "DB 暇なら 6", "6"],
            ["6", "Proxy ピン率", "多重化死亡", "接続は健全", "7"],
            ["7", "タスクを倍にして TPS が倍か", "倍なら水平可。倍でないなら共有天井", "天井の段を 2〜6 で再特定", "処方箋"],
            ["8", "Connect Time だけ長い", "TLS、Keep-Alive、枯渇、DNS", "アプリ本体ではない", "ジェネレータと接続設計"],
        ],
    )


def metrics(wb: Workbook) -> None:
    ws = sheet_base(wb, "18_監視メトリクス", TEAL)
    apply_widths(ws, {1: 18, 2: 34, 3: 22, 4: 40, 5: 36, 6: 18})
    banner(ws, 1, 6, "試験中に同時に見るメトリクス", NAVY)
    note(ws, 2, 6, "X 軸は UTC。注釈にランプ開始・定常・帳票オンを入れる。標準解像度 60s。短いスパイクは 20s またはアプリメトリクス。")
    headers(ws, 4, ["段", "メトリクス", "名前空間の目安", "読み方", "警報の芽", "必須"])
    rows = [
        ["JMeter", "TPS, p95, error, threads, connect", "JTL / Influx", "外測", "スレッド減少はジェネレータ死", "必須"],
        ["ジェネレータ", "CPU, MEM, NET, FD, GC", "OS", "70% 超は無効疑い", "GC 停止", "必須"],
        ["API GW", "Count, Latency, IntegrationLatency, 4XX, 5XX", "AWS/ApiGateway", "通すランだけ", "429, Latency差", "条件"],
        ["ALB", "RequestCount, TargetResponseTime, ELB 5xx, UnHealthy, RequestCountPerTarget, ActiveConnectionCount", "AWS/ApplicationELB", "アプリ外測の主", "502/503/504", "必須"],
        ["ECS Service", "CPUUtilization, MemoryUtilization, TaskCount", "AWS/ECS", "平均温度", "Desired≠Running", "必須"],
        ["ECS Task", "CpuUtilized, MemoryUtilized", "ContainerInsights / Logs", "偏り", "一部 100%", "必須"],
        ["アプリ JVM", "heap, GC pause, threads, pool wait", "micrometer 等", "待ちの中身", "GC 秒、pool wait", "必須"],
        ["ElastiCache", "HitRate, EngineCPU, Evictions, CurrConnections, ReplicationLag", "AWS/ElastiCache", "棚の状態", "EngineCPU 高、Evictions", "必須"],
        ["RDS Proxy", "ClientConnections, DatabaseConnections, CurrentlySessionPinned, Borrowed", "AWS/RDS", "多重化", "Pinned≈Borrowed", "Proxy 利用時必須"],
        ["Aurora", "ServerlessDatabaseCapacity, ACUUtilization, CPU, BufferCacheHitRatio, DatabaseConnections, CommitLatency, Deadlocks", "AWS/RDS", "床の広さと待ち", "ACU 上限、HitRatio 低下", "必須"],
        ["帳票", "queue, in-flight, heap, tmp files, job fail", "製品 / カスタム", "印刷室", "キュー右肩、504 同時", "帳票必須"],
        ["トレース", "span duration", "X-Ray / OTel", "コード行", "特定 SQL/HTTP", "推奨"],
    ]
    put_rows(ws, 5, rows)
    ws.auto_filter.ref = f"A4:F{4 + len(rows)}"


def passfail(wb: Workbook) -> None:
    ws = sheet_base(wb, "19_合否と失敗", GREEN)
    apply_widths(ws, {1: 28, 2: 28, 3: 28, 4: 28, 5: 40})
    banner(ws, 1, 5, "合否例（数値はプロジェクトで上書きする）", NAVY)
    note(
        ws,
        2,
        5,
        "前提例: 内部 ALB、タスク 8、ACU ウォーム済み、帳票 5% 同期、ウォームアップ 10 分除外、定常 30 分。値は例。計画書の NFR に置き換える。",
    )
    headers(ws, 4, ["項目", "合格", "条件付き（署名が要る）", "不合格", "測る場所"])
    put_rows(
        ws,
        5,
        [
            ["主要参照 p95", "≤ 300 ms", "300〜500 ms", "> 500 ms", "JMeter / TargetResponseTime"],
            ["更新 p95", "≤ 500 ms", "500〜800 ms", "> 800 ms", "同上"],
            ["帳票 p95", "≤ 8 s", "8〜15 s", "> 15 s または 504", "帳票ラベル"],
            ["真の失敗率", "< 0.1%", "0.1〜1%", "> 1%", "アサーション定義を明記"],
            ["目標 TPS", "定常維持", "5% 未満の欠落", "維持不能", "Transactions/s"],
            ["タスク CPU 平均", "< 70%", "70〜85%", "85% 超が持続", "ECS Service（余力判定）"],
            ["タスク偏り", "max-min < 20pt", "20〜40pt", "一部だけ 100%", "Task 次元"],
            ["Proxy ピン留め", "低位安定", "中程度", "貸出の大半", "Pinned / DB conn"],
            ["ACU", "上限に余裕", "時々接近", "張り付き＋p95 悪化", "ServerlessDatabaseCapacity"],
            ["キャッシュヒット", "本番同等", "やや低い", "コールドのまま", "HitRate"],
            ["ジェネレータ CPU", "< 70%", "70〜80%", "> 80%（試験無効）", "OS"],
            ["429（GW 通し時）", "0", "ごく少数", "持続", "API GW"],
            ["直線性（2倍タスク）", "TPS ≒ 2 倍", "1.5〜1.8 倍", "ほぼ増えない", "ランC拡張"],
        ],
    )


def failures(wb: Workbook) -> None:
    ws = sheet_base(wb, "19b_失敗の見分け", RED)
    apply_widths(ws, {1: 28, 2: 36, 3: 36, 4: 36, 5: 28})
    banner(ws, 1, 5, "症状 → 疑う段 → 確認 → 誤薬", NAVY)
    headers(ws, 3, ["症状", "疑う段", "確認", "誤薬", "正攻法"])
    put_rows(
        ws,
        4,
        [
            ["429 が増える", "API Gateway / WAF", "クォータ、Usage Plan、隣の試験", "ECS 増", "門を外すか上限申請、別ラン"],
            ["504 が 29s 付近", "API Gateway 統合", "Integration timeout、帳票時間", "ALB idle 延長だけ", "非同期化か GW 外し"],
            ["504 が 60s 付近", "ALB idle", "access log が idle と一致", "アプリ再起動", "idle 延長か進捗か非同期"],
            ["502", "タスク切断", "keep-alive、OOM、クラッシュ", "ALB 増", "graceful、タイムアウト整合"],
            ["503 ＋ Healthy 0", "起動・ヘルス", "新規タスク、重いヘルス", "コード最適化だけ", "ランプ、ヘルス軽量化"],
            ["503 ＋ Healthy あり急ランプ", "ALB スケール", "ConsumedLCU、短ランプ", "アプリ解析", "ランプ、必要なら事前相談"],
            ["CPU 低なのに遅い", "待ち（DB/帳票/外部）", "トレース、プール wait", "タスク CPU 増", "待ちの段を直す"],
            ["CPU 高で飽和", "アプリ計算", "1 タスク曲線", "DB 増", "最適化かタスクサイズ/台数"],
            ["平均 CPU 低・p99 高", "偏り", "TaskId 別、sticky", "全体増設", "偏りの原因"],
            ["タスク倍でも TPS 増えない", "共有天井", "帳票、Proxy、Aurora、キャッシュ", "さらに ECS 増", "天井の段"],
            ["p95 が階段、少し遅れ ACU 上昇", "Aurora 伸縮待ち", "ServerlessDatabaseCapacity", "SQL 全面見直しだけ", "min 引き上げ、固定ラン比較"],
            ["初回だけ十数秒", "ポーズ/コールド", "min ACU、バッファ", "アプリバグ", "ポーズ禁止、ウォーム"],
            ["接続エラー、max_connections", "プール×タスク", "Proxy ありなし、ピン率", "タイムアウト延長", "プール再計算"],
            ["Pinned ≈ Borrowed", "ピン留め", "16KB SQL、SET", "Proxy 台数増", "SQL/初期化見直しか廃止"],
            ["HitRate 低＋DB 炎上", "キャッシュミス/stampede", "TTL、キー、コールド", "Aurora max 無限", "ウォーム、単一フライト"],
            ["EngineCPU 高・HitRate 高", "ホットキー", "同一キー", "ノード縦増だけ", "分割、レプリカ、ローカルキャッシュ"],
            ["帳票オンで参照 API まで遅い", "スレッド飢餓", "tomcat threads、帳票 in-flight", "ECS 増（印刷室は 1）", "隔離・非同期・並列上限"],
            ["Connect Time だけ増", "TLS/枯渇/DNS", "Keep-Alive、FD、ジェネレータ CPU", "アプリチューニング", "接続再利用、分散"],
            ["後半だけ 401", "トークン期限", "JWT exp", "性能劣化", "更新フロー"],
            ["後半だけ unique 違反", "データ設計", "CSV Recycle", "DB 性能", "データ量と Recycle 方針"],
        ],
    )


def checklist(wb: Workbook) -> None:
    ws = sheet_base(wb, "20_チェックリスト", GREEN)
    apply_widths(ws, {1: 14, 2: 8, 3: 70, 4: 22, 5: 22, 6: 18})
    banner(ws, 1, 6, "実行チェックリスト（箱を埋めてから打つ）", NAVY)
    note(ws, 2, 6, "ステータス列は未了 / 完了 / 該当なし。フィルタして未了だけ残す。")
    headers(ws, 4, ["工程", "状態", "項目", "証拠の置き場", "担当", "日付"])
    items = [
        ["計画", "未了", "目的が一文である", "計画書", "", ""],
        ["計画", "未了", "打ち込み位置がランごとに書かれている", "09 相当表", "", ""],
        ["計画", "未了", "API Gateway を通す/通さない理由がある", "計画書", "", ""],
        ["計画", "未了", "1 タスク飽和のランがある", "ランC", "", ""],
        ["計画", "未了", "帳票あり/なしが分かれている", "ランB/B'/E", "", ""],
        ["計画", "未了", "合否が数値である（p95、エラー定義、除外）", "19 相当", "", ""],
        ["計画", "未了", "データと個人情報が承認済み", "チケット", "", ""],
        ["計画", "未了", "スタブと課金上限がある", "計画書", "", ""],
        ["計画", "未了", "オートスケールを止めるランと動かすランがある", "ラン表", "", ""],
        ["計画", "未了", "Aurora を試験中ポーズさせない（min>0、ウォーム）", "パラメータ", "", ""],
        ["計画", "未了", "監視とログ保持が延びている", "AWS コンソール", "", ""],
        ["計画", "未了", "Service Quotas（GW RPS 等）を日付つきで写した", "計画書", "", ""],
        ["JMeter", "未了", "5.6.x、Java 17 推奨", "java -version", "", ""],
        ["JMeter", "未了", "非 GUI コマンドが確定", "22_コマンド", "", ""],
        ["JMeter", "未了", "Tree / Table リスナーオフ", "JMX", "", ""],
        ["JMeter", "未了", "CSV JTL、必要列だけ", "user.properties", "", ""],
        ["JMeter", "未了", "Cookie Manager、Keep-Alive 本番相当", "JMX", "", ""],
        ["JMeter", "未了", "思考時間あり", "Timer", "", ""],
        ["JMeter", "未了", "相関がデバッグ負荷で 100% 成功", "Tree（デバッグ時のみ）", "", ""],
        ["JMeter", "未了", "Groovy は cache + vars.get", "JSR223", "", ""],
        ["JMeter", "未了", "ホストとスレッドは -J で外部化", "CLI", "", ""],
        ["JMeter", "未了", "ジェネレータ監視がある", "OS メトリクス", "", ""],
        ["実行", "未了", "NTP / UTC", "timedatectl 等", "", ""],
        ["実行", "未了", "ウォームアップを統計から除外する印", "報告書", "", ""],
        ["実行", "未了", "開始終了 UTC を記録", "実行ログ", "", ""],
        ["実行", "未了", "429/502/503/504 の定義を共有", "チャット/計画", "", ""],
        ["実行", "未了", "隣の試験とアカウント上限を奪い合わない", "予定表", "", ""],
        ["実行", "未了", "Desired タスク数とオートスケール状態を記録", "ECS 画面", "", ""],
        ["事後", "未了", "HTML + JTL + ログ + CloudWatch を一式保管", "成果物フォルダ", "", ""],
        ["事後", "未了", "1 ページ結論（合格/条件付き/不合格）", "報告書", "", ""],
        ["事後", "未了", "ボトルネック段が特定されている", "17 の結果", "", ""],
        ["事後", "未了", "次アクションの担当がいる", "チケット", "", ""],
        ["事後", "未了", "再現コマンドが残っている", "README/報告書", "", ""],
    ]
    put_rows(ws, 5, items)
    ws.auto_filter.ref = f"A4:F{4 + len(items)}"
    dv = DataValidation(type="list", formula1='"未了,完了,該当なし"', allow_blank=True)
    dv.error = "未了 / 完了 / 該当なし から選択"
    dv.errorTitle = "状態"
    dv.prompt = "状態を選択"
    dv.promptTitle = "チェック"
    ws.add_data_validation(dv)
    dv.add(f"B5:B{4 + len(items)}")
    # conditional colors for 状態
    ws.conditional_formatting.add(
        f"B5:B{4 + len(items)}",
        FormulaRule(formula=['B5="完了"'], fill=fill(PALE_GREEN)),
    )
    ws.conditional_formatting.add(
        f"B5:B{4 + len(items)}",
        FormulaRule(formula=['B5="未了"'], fill=fill(PALE_RED)),
    )


def analogies(wb: Workbook) -> None:
    ws = sheet_base(wb, "21_たとえ話集", GOLD)
    apply_widths(ws, {1: 22, 2: 40, 3: 50, 4: 40})
    banner(ws, 1, 4, "小学生（と経営層）に説明するたとえ", NAVY)
    note(ws, 2, 4, "たとえは入口。報告書の数値はたとえで置き換えない。")
    headers(ws, 4, ["対象", "たとえ", "正確に対応すること", "説明の締め"])
    put_rows(
        ws,
        5,
        [
            ["JMeter", "お客役をたくさん作る機械", "ブラウザではない。注文（HTTP）だけ出す", "テレビ（GUI）を付けたまま走ると機械が先に疲れる"],
            ["スレッド", "お客ひとり", "同時に厨房にいる人数", "人数≠1秒の皿数"],
            ["ランプ", "少しずつ校門を開ける", "オートスケールを起こす時間", "いきなり全員入れるのは別ゲーム"],
            ["p95", "100人中95人目の待ち", "遅い人の側", "平均は足の速い人に引っ張られる"],
            ["API Gateway", "校門の警備と回転木戸", "RPS とバーストと学生証", "木戸の幅を厨房の速さと呼ばない"],
            ["429", "満員で門の外に並ばされる", "トークンが空", "料理人を増やしても木戸は広がらない"],
            ["29秒", "門は29秒で追い出す", "統合タイムアウト", "印刷が30秒かかる成績表は門の外で失敗する"],
            ["ALB", "配膳口", "空いている窓へ", "閉じた窓に渡すと 502"],
            ["ECS タスク", "調理員ひとり", "CPU とメモリの一人分", "クラス平均点で火事は分からない"],
            ["1 タスク試験", "徒競走の個人タイム", "単位能力", "4人リレーが速くならないならバトン（DB）"],
            ["ElastiCache", "手元の棚", "ヒットとホットキー", "棚が空だと全員が倉庫へ走る"],
            ["RDS Proxy", "整理券機", "席の使い回し", "専用ロッカーを使うと席が固定（ピン留め）"],
            ["Aurora ACU", "伸縮する床", "約 2GiB ごとの広さ", "広げる時間は別に測る。畳んだ倉庫は起こす時間が要る"],
            ["帳票", "印刷室", "同時に何台刷れるか", "料理人を増やしてもプリンタは増えない"],
            ["ソーク", "何時間も同じ給食", "リーク", "最初の10分が速くても、夕方に床がゴミだらけなら不合格"],
            ["ウォームアップ", "準備運動", "キャッシュと JIT と ACU", "準備運動のタイムを記録に残さない"],
        ],
    )


def commands(wb: Workbook) -> None:
    ws = sheet_base(wb, "22_コマンド集", SLATE)
    apply_widths(ws, {1: 22, 2: 78, 3: 40, 4: 22})
    banner(ws, 1, 4, "オペレーション用コマンド（Windows / 一般）", NAVY)
    note(ws, 2, 4, "パスは環境で置き換える。Heap は set HEAP=-Xms4g -Xmx4g など。大きすぎる Xmx は GC 停止を長くする。")
    headers(ws, 4, ["用途", "コマンド / クエリ", "説明", "注意"])
    put_rows(
        ws,
        5,
        [
            ["本試験＋HTML", "jmeter -n -t test.jmx -l results.jtl -e -o html-report -j jmeter.log", "非 GUI 標準", "html-report は空ディレクトリ"],
            ["プロパティ差し込み", "jmeter -n -t test.jmx -l results.jtl -e -o html-report -Jthreads=80 -Jhost=internal-alb.example.local -Jduration=1800", "スレッド・ホスト・秒", "JMX 側は ${__P(threads,10)}"],
            ["追加プロパティ", "jmeter -n -t test.jmx -q extra.properties -l results.jtl", "セットで切替", "jmeter.properties は直接編集しない"],
            ["既存 JTL から HTML", "jmeter -g results.jtl -o html-report", "再集計", "粒度と APDEX は user.properties"],
            ["Java 確認", "java -version", "17 推奨", "8 未満は 5.6.3 不可"],
            ["ECS タスク CPU (Logs Insights)", 'fields @timestamp, TaskId, CpuUtilized, MemoryUtilized, CpuReserved | filter Type = "Task" | stats avg(CpuUtilized), max(CpuUtilized), avg(MemoryUtilized), max(MemoryUtilized) by TaskId', "偏り", "カスタムメトリクス化は課金"],
            ["ALB idle 確認", 'aws elbv2 describe-load-balancer-attributes --load-balancer-arn ARN --query "Attributes[?Key==\'idle_timeout.timeout_seconds\']"', "60s 既定", "帳票同期と突合"],
            ["GW クォータ", "Service Quotas コンソールで API Gateway Throttle rate", "10,000 既定", "リージョン差"],
            ["結果列（JTL）", "timeStamp,elapsed,label,responseCode,success,bytes,Latency,Connect,allThreads,...", "CSV ヘッダ", "XML にしない"],
        ],
    )


def references(wb: Workbook) -> None:
    ws = sheet_base(wb, "23_参考情報", NAVY)
    apply_widths(ws, {1: 28, 2: 78, 3: 44})
    banner(ws, 1, 3, "一次情報（数値は試験前に再確認）", NAVY)
    note(ws, 2, 3, "本ブックの数値は 2026-09 時点の公式ドキュメントに基づく。クォータ・ACU 上限・タイムアウトはアカウントとエンジン版で変わる。")
    headers(ws, 4, ["資料", "URL / 場所", "本ブックでの使い方"])
    put_rows(
        ws,
        5,
        [
            ["JMeter 本体・最新 5.6.3", "https://jmeter.apache.org/  / download_jmeter.cgi", "版と Java 要件"],
            ["JMeter Best Practices", "https://jmeter.apache.org/usermanual/best-practices.html", "非 GUI、リスナー、Groovy"],
            ["JMeter Dashboard", "https://jmeter.apache.org/usermanual/generating-dashboard.html", "APDEX、粒度 >=1s"],
            ["JMeter Changes 5.6.3", "https://jmeter.apache.org/changes.html", "既知修正"],
            ["API Gateway クォータ", "https://docs.aws.amazon.com/apigateway/latest/developerguide/limits.html", "10k RPS、バースト、29s"],
            ["API Gateway throttling", "https://docs.aws.amazon.com/apigateway/latest/developerguide/http-api-throttling.html", "トークンバケツ"],
            ["ALB idle timeout", "https://docs.aws.amazon.com/elasticloadbalancing/latest/application/edit-load-balancer-attributes.html", "既定 60s、1-4000、keep-alive 整合"],
            ["ALB TargetResponseTime", "https://repost.aws/knowledge-center/alb-troubleshoot-targetresponsetime", "target_processing_time"],
            ["ECS サービス使用率", "https://docs.aws.amazon.com/AmazonECS/latest/developerguide/service_utilization.html", "予約対使用"],
            ["Container Insights", "https://docs.aws.amazon.com/AmazonCloudWatch/latest/monitoring/ContainerInsights.html", "enhanced 2024-12-02"],
            ["ECS 高解像度メトリクス", "https://aws.amazon.com/blogs/aws/amazon-ecs-introduces-new-high-resolution-metrics-for-faster-service-auto-scaling/", "20s、2026-06"],
            ["Aurora Sv2 容量", "https://docs.aws.amazon.com/AmazonRDS/latest/AuroraUserGuide/aurora-serverless-v2.setting-capacity.html", "min/max、バッファ"],
            ["Aurora オートポーズ", "https://docs.aws.amazon.com/AmazonRDS/latest/AuroraUserGuide/aurora-serverless-v2-auto-pause.html", "min=0"],
            ["Aurora 高速スケール（ブログ）", "https://aws.amazon.com/blogs/database/faster-scaling-for-aurora-serverless-to-support-agentic-ai-and-other-spiky-workloads/", "+12 ACU/s 等（発表値）"],
            ["RDS Proxy ピン留め", "https://docs.aws.amazon.com/AmazonRDS/latest/UserGuide/rds-proxy-pinning.html", "16KB、SET"],
            ["RDS Proxy 接続", "https://docs.aws.amazon.com/AmazonRDS/latest/UserGuide/rds-proxy-connections.html", "24h、idle"],
            ["RDS Proxy ダッシュボード", "https://docs.aws.amazon.com/AmazonRDS/latest/UserGuide/rds-proxy-monitoring-dashboard.html", "30% ヘッドルーム"],
            ["SVF 並列・ログ", "ウイングアーク KB（UCX 並列、一時ディレクトリ、デバッグ OFF）", "帳票天井"],
            ["JasperReports 負荷", "Jaspersoft Community（JMeter での負荷、virtualizer、同時実行）", "メモリが先に限界になりやすい"],
            ["同梱 Markdown", "docs/JMeter_全体パフォーマンステスト完全解説.md", "文章の深掘り本体"],
        ],
    )


if __name__ == "__main__":
    build()
