"""channel_list.csv に記載された全チャンネルの直近1ヶ月分のメッセージ
(スレッド返信含む) を取得して、1つのテキストファイルに書き出すスクリプト。

使い方:
    python fetch_channels_history.py [--days 30] [--csv channel_list.csv] [--out 出力先.txt]

事前準備:
    .env もしくは環境変数に SLACK_BOT_TOKEN を設定しておくこと。
    Bot は対象チャンネルに招待されている必要がある。
"""

from __future__ import annotations

import argparse
import csv
import os
import sys
import time
from datetime import datetime, timedelta
from typing import Any, Dict, List, Optional

import requests

CONVERSATIONS_HISTORY_URL = "https://slack.com/api/conversations.history"
CONVERSATIONS_REPLIES_URL = "https://slack.com/api/conversations.replies"
CONVERSATIONS_INFO_URL = "https://slack.com/api/conversations.info"
USERS_INFO_URL = "https://slack.com/api/users.info"


# ---------------------------------------------------------------------------
# 環境変数 / トークン
# ---------------------------------------------------------------------------
def load_dotenv(path: str = ".env") -> None:
    """簡易 .env ローダー (make_melmaga.py と同等)。"""
    if not os.path.exists(path):
        return
    try:
        with open(path, "r", encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if not line or line.startswith("#") or "=" not in line:
                    continue
                key, val = line.split("=", 1)
                key = key.strip()
                val = val.strip().strip('"').strip("'")
                if key and key not in os.environ:
                    os.environ[key] = val
    except Exception:
        pass


def get_token() -> str:
    token = os.getenv("SLACK_BOT_TOKEN")
    if not token:
        raise RuntimeError("環境変数 SLACK_BOT_TOKEN が設定されていません。")
    return token


# ---------------------------------------------------------------------------
# HTTP ヘルパー (リトライ・レートリミット対応)
# ---------------------------------------------------------------------------
def http_get_json(
    url: str,
    token: str,
    params: Dict[str, Any],
    max_retries: int = 5,
) -> Dict[str, Any]:
    headers = {"Authorization": f"Bearer {token}"}
    for attempt in range(max_retries):
        res = requests.get(url, headers=headers, params=params, timeout=30)
        if res.status_code == 429:
            retry_after = int(res.headers.get("Retry-After", "1"))
            time.sleep(max(retry_after, 1))
            continue
        res.raise_for_status()
        data = res.json()
        if data.get("ok"):
            return data
        err = data.get("error")
        if err == "ratelimited":
            time.sleep(2 ** attempt)
            continue
        raise RuntimeError(f"Slack API error ({url}): {data}")
    raise RuntimeError(f"リトライ上限を超えました: {url}")


# ---------------------------------------------------------------------------
# チャンネルリスト読み込み
# ---------------------------------------------------------------------------
def load_channels(csv_path: str) -> List[Dict[str, str]]:
    channels: List[Dict[str, str]] = []
    with open(csv_path, "r", encoding="utf-8", errors="replace", newline="") as f:
        reader = csv.DictReader(f)
        for row in reader:
            cid = (row.get("channel_id") or "").strip()
            if not cid:
                continue
            channels.append(
                {
                    "channel_id": cid,
                    "channel_name": (row.get("channel_name") or "").strip(),
                }
            )
    return channels


# ---------------------------------------------------------------------------
# Slack API ラッパー
# ---------------------------------------------------------------------------
def fetch_channel_info(token: str, channel_id: str) -> Optional[Dict[str, Any]]:
    try:
        data = http_get_json(CONVERSATIONS_INFO_URL, token, {"channel": channel_id})
        return data.get("channel")
    except Exception as e:
        print(f"  conversations.info 失敗 ({channel_id}): {e}", file=sys.stderr)
        return None


_user_cache: Dict[str, str] = {}


def get_user_name(token: str, user_id: str) -> str:
    if not user_id:
        return ""
    if user_id in _user_cache:
        return _user_cache[user_id]
    try:
        data = http_get_json(USERS_INFO_URL, token, {"user": user_id})
        user = data.get("user", {}) or {}
        profile = user.get("profile", {}) or {}
        name = (
            profile.get("display_name")
            or profile.get("real_name")
            or user.get("real_name")
            or user.get("name")
            or user_id
        )
    except Exception:
        name = user_id
    _user_cache[user_id] = name
    return name


def fetch_history(token: str, channel_id: str, oldest_ts: float) -> List[Dict[str, Any]]:
    """指定チャンネルの oldest_ts 以降のメッセージを全件取得 (ページング対応)。"""
    messages: List[Dict[str, Any]] = []
    cursor: Optional[str] = None
    while True:
        params: Dict[str, Any] = {
            "channel": channel_id,
            "oldest": f"{oldest_ts:.6f}",
            "limit": 200,
        }
        if cursor:
            params["cursor"] = cursor
        data = http_get_json(CONVERSATIONS_HISTORY_URL, token, params)
        messages.extend(data.get("messages", []))
        cursor = (data.get("response_metadata") or {}).get("next_cursor") or None
        if not cursor:
            break
        time.sleep(0.3)
    return messages


def fetch_replies(
    token: str, channel_id: str, thread_ts: str, oldest_ts: float
) -> List[Dict[str, Any]]:
    """スレッドの返信を取得。期間内のものだけ返す。"""
    messages: List[Dict[str, Any]] = []
    cursor: Optional[str] = None
    while True:
        params: Dict[str, Any] = {
            "channel": channel_id,
            "ts": thread_ts,
            "limit": 200,
        }
        if cursor:
            params["cursor"] = cursor
        data = http_get_json(CONVERSATIONS_REPLIES_URL, token, params)
        for m in data.get("messages", []):
            try:
                if float(m.get("ts", "0")) >= oldest_ts:
                    messages.append(m)
            except Exception:
                continue
        cursor = (data.get("response_metadata") or {}).get("next_cursor") or None
        if not cursor:
            break
        time.sleep(0.3)
    return messages


# ---------------------------------------------------------------------------
# 整形
# ---------------------------------------------------------------------------
def format_ts(ts_str: str) -> str:
    try:
        return datetime.fromtimestamp(float(ts_str)).strftime("%Y-%m-%d %H:%M:%S")
    except Exception:
        return ts_str or ""


def format_message(token: str, m: Dict[str, Any], indent: str = "") -> str:
    if m.get("user"):
        speaker = get_user_name(token, m["user"])
    else:
        speaker = m.get("username") or m.get("bot_id") or "bot"
    ts = format_ts(m.get("ts", ""))
    text = (m.get("text") or "").rstrip()

    lines = [f"{indent}[{ts}] {speaker}:"]
    if text:
        for line in text.splitlines():
            lines.append(f"{indent}  {line}")
    else:
        lines.append(f"{indent}  (本文なし)")

    for f in m.get("files") or []:
        name = f.get("name") or f.get("title") or ""
        url = f.get("url_private") or f.get("permalink") or ""
        lines.append(f"{indent}  [添付] {name} {url}".rstrip())

    for att in m.get("attachments") or []:
        title = att.get("title") or ""
        text_a = att.get("text") or att.get("fallback") or ""
        if title or text_a:
            lines.append(f"{indent}  [attachment] {title} {text_a}".rstrip())

    return "\n".join(lines)


# ---------------------------------------------------------------------------
# メイン処理
# ---------------------------------------------------------------------------
def months_ago(base: datetime, months: int) -> datetime:
    """base から months ヶ月前の datetime をカレンダーベースで返す。"""
    y = base.year
    m = base.month - months
    while m <= 0:
        m += 12
        y -= 1
    # 月末日が存在しない場合は前月末に丸める (例: 3/31 -> 1ヶ月前 -> 2/28 or 29)
    day = base.day
    while True:
        try:
            return base.replace(year=y, month=m, day=day)
        except ValueError:
            day -= 1
            if day < 1:
                raise


def parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(description="Slack チャンネルの過去メッセージをテキスト出力")
    p.add_argument("--csv", default="channel_list.csv", help="チャンネルリストCSVのパス")
    p.add_argument(
        "--months",
        type=int,
        default=1,
        help="さかのぼる月数 (デフォルト: 1ヶ月)。--days 指定時はそちらが優先。",
    )
    p.add_argument(
        "--days",
        type=int,
        default=None,
        help="さかのぼる日数。指定時は --months より優先される。",
    )
    p.add_argument("--out", default=None, help="出力先ファイルパス")
    return p.parse_args()


def main() -> None:
    args = parse_args()
    load_dotenv()
    token = get_token()

    channels = load_channels(args.csv)
    if not channels:
        raise RuntimeError(f"{args.csv} からチャンネルを読み込めませんでした。")

    now = datetime.now()
    if args.days is not None:
        oldest_dt = now - timedelta(days=args.days)
        span_label = f"直近{args.days}日"
    else:
        oldest_dt = months_ago(now, args.months)
        span_label = f"直近{args.months}ヶ月"
    oldest_ts = oldest_dt.timestamp()

    out_path = args.out or f"slack_export_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"

    print(f"対象チャンネル数: {len(channels)}", file=sys.stderr)
    print(f"期間: {oldest_dt.strftime('%Y-%m-%d %H:%M:%S')} ～ 現在 ({span_label})", file=sys.stderr)
    print(f"出力先: {out_path}", file=sys.stderr)

    total_msgs = 0

    with open(out_path, "w", encoding="utf-8") as out:
        out.write("# Slack 過去メッセージ エクスポート\n")
        out.write(f"# 生成日時 : {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
        out.write(f"# 対象期間 : {oldest_dt.strftime('%Y-%m-%d %H:%M:%S')} 以降 ({span_label})\n")
        out.write(f"# チャンネル数 : {len(channels)}\n\n")

        for idx, ch in enumerate(channels, 1):
            cid = ch["channel_id"]
            info = fetch_channel_info(token, cid)
            name = (info or {}).get("name") or ch.get("channel_name") or cid
            print(f"[{idx}/{len(channels)}] #{name} ({cid}) を取得中...", file=sys.stderr)

            out.write("=" * 80 + "\n")
            out.write(f"## #{name}  ({cid})\n")
            out.write("=" * 80 + "\n\n")

            try:
                messages = fetch_history(token, cid, oldest_ts)
            except Exception as e:
                print(f"  history 取得失敗: {e}", file=sys.stderr)
                out.write(f"(取得失敗: {e})\n\n")
                continue

            messages.sort(key=lambda m: float(m.get("ts", "0")))

            if not messages:
                out.write("(期間内のメッセージなし)\n\n")
                continue

            ch_msg_count = 0
            for m in messages:
                out.write(format_message(token, m) + "\n")
                ch_msg_count += 1

                if m.get("reply_count") and m.get("thread_ts") == m.get("ts"):
                    try:
                        replies = fetch_replies(token, cid, m["ts"], oldest_ts)
                    except Exception as e:
                        out.write(f"    (返信取得失敗: {e})\n")
                        replies = []
                    for r in replies:
                        if r.get("ts") == m.get("ts"):
                            continue
                        out.write(format_message(token, r, indent="    └ ") + "\n")
                        ch_msg_count += 1
                out.write("\n")

            out.write(f"-- メッセージ数: {ch_msg_count} --\n\n")
            total_msgs += ch_msg_count
            time.sleep(0.5)

        out.write("\n")
        out.write(f"# 合計メッセージ数: {total_msgs}\n")

    print(f"完了: {out_path} (合計 {total_msgs} 件)", file=sys.stderr)


if __name__ == "__main__":
    main()
