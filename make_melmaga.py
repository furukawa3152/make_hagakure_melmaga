"""メルマガ HTML を組み立てるユーティリティ。

使い方:
    from make_melmaga import get_issue_vars, apply_template_vars, build_newsletter_html

    vars = get_issue_vars()          # {"year": "2026", "month": "5"}
    html = build_newsletter_html()   # template.txt + mail_body.html を結合
"""

from __future__ import annotations

from datetime import datetime
from pathlib import Path
from typing import Dict, Optional
from zoneinfo import ZoneInfo

JST = ZoneInfo("Asia/Tokyo")

DEFAULT_TEMPLATE = Path("template.txt")
DEFAULT_MAIL_BODY = Path("mail_body.html")


def get_issue_vars(when: Optional[datetime] = None) -> Dict[str, str]:
    """${year}年${month}月号 用の変数を返す。

    Args:
        when: 号数の基準日時。省略時は現在（Asia/Tokyo）。

    Returns:
        {"year": "2026", "month": "5"} のような dict（いずれも str）。
    """
    dt = when or datetime.now(JST)
    if dt.tzinfo is None:
        dt = dt.replace(tzinfo=JST)
    else:
        dt = dt.astimezone(JST)
    return {
        "year": str(dt.year),
        "month": str(dt.month),
    }


def apply_template_vars(template: str, variables: Dict[str, str]) -> str:
    """テンプレート内の ${key} を variables の値で置換する。"""
    result = template
    for key, value in variables.items():
        result = result.replace(f"${{{key}}}", value)
    return result


def read_text(path: Path) -> str:
    return path.read_text(encoding="utf-8")


def build_newsletter_html(
    template_path: Path = DEFAULT_TEMPLATE,
    mail_body_path: Path = DEFAULT_MAIL_BODY,
    when: Optional[datetime] = None,
    extra_vars: Optional[Dict[str, str]] = None,
) -> str:
    """template.txt と mail_body.html を結合して完成 HTML を返す。"""
    variables = get_issue_vars(when)
    variables["mailBody"] = read_text(mail_body_path)
    if extra_vars:
        variables.update(extra_vars)
    return apply_template_vars(read_text(template_path), variables)


if __name__ == "__main__":
    import argparse

    p = argparse.ArgumentParser(description="メルマガ HTML をプレビュー生成")
    p.add_argument("--template", type=Path, default=DEFAULT_TEMPLATE)
    p.add_argument("--body", type=Path, default=DEFAULT_MAIL_BODY)
    p.add_argument("--out", type=Path, default=Path("newsletter_preview.html"))
    p.add_argument("--year", type=int, help="号数の年（省略時は今月）")
    p.add_argument("--month", type=int, help="号数の月（省略時は今月）")
    args = p.parse_args()

    when: Optional[datetime] = None
    if args.year is not None or args.month is not None:
        now = datetime.now(JST)
        when = datetime(
            args.year or now.year,
            args.month or now.month,
            1,
            tzinfo=JST,
        )

    html = build_newsletter_html(args.template, args.body, when=when)
    args.out.write_text(html, encoding="utf-8")
    v = get_issue_vars(when)
    print(f"生成完了: {args.out}  ({v['year']}年{v['month']}月号)")
