#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""PDFファイル名ルール管理および自動仕分けツール.

仕様書: ファイル移動仕様書.md
設計原則: 基本設計.md

Python標準ライブラリのみで動作する。外部ネットワークへは一切送信しない。

  py pdf_sorter_app.py server    ルール編集画面を起動
  py pdf_sorter_app.py sort      rules.json に従いPDF移動を実行
  py pdf_sorter_app.py preview   rules.json に従い移動予定先をコンソール表示
"""

from __future__ import annotations

import csv
import json
import os
import re
import shutil
import sys
import threading
import time
import webbrowser
from datetime import datetime
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path

APP_NAME = "PDF自動仕分けツール"
APP_VERSION = "1.0.0"

HOST = "127.0.0.1"
PORT = 8765
PORT_RETRY = 10

RULES_FILE_NAME = "rules.json"
LOCK_FILE_NAME = ".sorter.lock"
LOG_FILE_NAME = "move_log.csv"
LAST_RUN_FILE_NAME = "last_run.json"
RULES_BACKUP_FOLDER = "_rules_backup"
RULES_BACKUP_KEEP = 20

MONTH_UNKNOWN_FOLDER = "_年月不明"
MONTH_AMBIGUOUS_FOLDER = "_年月確認要"

TARGET_EXTENSION = ".pdf"
EXCLUDE_PREFIXES = ("~$", ".")
EXCLUDE_SUFFIXES = (".tmp", ".part", ".partial", ".crdownload", ".filepart")

WINDOWS_FORBIDDEN_CHARS = '<>:"|?*'
WINDOWS_RESERVED_NAMES = {
    "CON", "PRN", "AUX", "NUL",
    *(f"COM{i}" for i in range(1, 10)),
    *(f"LPT{i}" for i in range(1, 10)),
}

UNSET_BASE_PLACEHOLDER = "{BASE_FOLDER}"

# --- 結果コード -----------------------------------------------------------
# 実行時
R_MOVED = "MOVED"
R_COPIED_AND_DELETED = "COPIED_AND_DELETED"
R_COPIED = "COPIED"
R_DUPLICATE_EVACUATED = "DUPLICATE_EVACUATED"
R_UNKNOWN = "UNKNOWN"
R_MONTH_UNKNOWN = "MONTH_UNKNOWN"
R_MONTH_AMBIGUOUS = "MONTH_AMBIGUOUS"
R_SKIP = "SKIP"
R_SKIP_FILE_NOT_STABLE = "SKIP_FILE_NOT_STABLE"
R_UNDONE = "UNDONE"
# プレビュー時
R_MATCH = "MATCH"
R_DUPLICATE_WILL_EVACUATE = "DUPLICATE_WILL_EVACUATE"
# エラー
E_CREATE_FOLDER = "ERROR_CREATE_FOLDER"
E_MOVE_FAILED = "ERROR_MOVE_FAILED"
E_COPY_FAILED = "ERROR_COPY_FAILED"
E_DELETE_SOURCE_FAILED = "ERROR_DELETE_SOURCE_FAILED"
E_VERIFY_FAILED = "ERROR_VERIFY_FAILED"
E_PATH_TOO_LONG = "ERROR_PATH_TOO_LONG"
E_RULE_INVALID = "ERROR_RULE_INVALID"
E_LOCKED = "ERROR_LOCKED"
E_RULES_JSON_INVALID = "ERROR_RULES_JSON_INVALID"
E_FORBIDDEN_PATH = "ERROR_FORBIDDEN_PATH"
E_UNDO_FAILED = "ERROR_UNDO_FAILED"

DEFAULT_SETTINGS = {
    "common_base": UNSET_BASE_PLACEHOLDER,
    "unknown_folder": "_unknown",
    "log_folder": "_log",
    "use_month_folder": True,
    "month_folder_format": "YYYY-MM",
    "month_source": "filename",
    "month_fallback": "unknown_month",
    "duplicate_mode": "evacuate",
    "evacuation_folder_prefix": "避難用",
    "forbid_overwrite": True,
    "move_strategy": "copy_verify_delete",
    "verify_after_move": True,
    "check_file_stable": True,
    "min_file_age_seconds": 10,
    "backup_rules_on_save": True,
    "max_path_length": 240,
    "rules": [],
}

DEFAULT_RULE = {
    "enabled": True,
    "priority": 1,
    "name": "",
    "extension": TARGET_EXTENSION,
    "contains_all": [],
    "contains_any": [],
    "not_contains": [],
    "destination_base": "",
    "destination_subfolder": "",
}

ENUMS = {
    "month_source": ("filename",),
    "month_fallback": ("current_month", "unknown_month", "error"),
    "duplicate_mode": ("evacuate", "skip"),
    "move_strategy": ("move", "copy_verify_delete", "copy_only"),
    "month_folder_format": ("YYYY-MM", "YYYY_MM", "YYYYMM"),
}


# =========================================================================
# 基本ユーティリティ
# =========================================================================

def app_dir() -> Path:
    """スクリプトが置かれているフォルダ（＝仕分け対象フォルダ）。"""
    return Path(__file__).resolve().parent


def rules_path() -> Path:
    return app_dir() / RULES_FILE_NAME


def lock_path() -> Path:
    return app_dir() / LOCK_FILE_NAME


def log_dir(settings: dict) -> Path:
    return app_dir() / settings.get("log_folder", DEFAULT_SETTINGS["log_folder"])


def log_path(settings: dict) -> Path:
    return log_dir(settings) / LOG_FILE_NAME


def last_run_path(settings: dict) -> Path:
    return log_dir(settings) / LAST_RUN_FILE_NAME


def now_stamp() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def setup_console() -> None:
    """Windowsコンソールでも日本語が化けないようにする。"""
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8", errors="replace")
        except (AttributeError, ValueError):
            pass


def echo(text: str = "") -> None:
    print(text, flush=True)


# =========================================================================
# 設定の検証と正規化
# =========================================================================

def _as_bool(value, field: str, errors: list) -> bool:
    if isinstance(value, bool):
        return value
    errors.append(f"{field} は true / false で指定してください（現在の値: {value!r}）")
    return bool(DEFAULT_SETTINGS.get(field, False))


def _as_int(value, field: str, errors: list, minimum: int, maximum: int) -> int:
    if isinstance(value, bool) or not isinstance(value, (int, float)):
        errors.append(f"{field} は数値で指定してください（現在の値: {value!r}）")
        return int(DEFAULT_SETTINGS.get(field, 0))
    number = int(value)
    if number < minimum or number > maximum:
        errors.append(f"{field} は {minimum} 以上 {maximum} 以下で指定してください（現在の値: {number}）")
        return int(DEFAULT_SETTINGS.get(field, minimum))
    return number


def _as_str(value, field: str, errors: list) -> str:
    if not isinstance(value, str):
        errors.append(f"{field} は文字列で指定してください（現在の値: {value!r}）")
        return str(DEFAULT_SETTINGS.get(field, ""))
    return value.strip()


def _as_str_list(value, field: str, errors: list) -> list:
    if value is None:
        return []
    if not isinstance(value, list):
        errors.append(f"{field} は文字列の配列で指定してください")
        return []
    items = []
    for entry in value:
        if not isinstance(entry, str):
            errors.append(f"{field} には文字列のみ指定できます（{entry!r} は使用できません）")
            continue
        text = entry.strip()
        if text:
            items.append(text)
    return items


def folder_name_error(name: str, label: str) -> str:
    """単一フォルダ名として使えるかを判定する（仕様書 14.2）。"""
    if not name:
        return f"{label}を入力してください"
    for char in WINDOWS_FORBIDDEN_CHARS + "/\\":
        if char in name:
            return f"{label}に使用できない文字 {char} が含まれています"
    return _segment_error(name, label)


def _segment_error(segment: str, label: str) -> str:
    if segment in (".", ".."):
        return f"{label}に {segment} は使用できません"
    if segment != segment.rstrip(" ."):
        return f"{label}の末尾に空白またはピリオドは使用できません"
    if segment.split(".")[0].upper() in WINDOWS_RESERVED_NAMES:
        return f"{label}に Windows の予約語 {segment} は使用できません"
    return ""


def subfolder_error(value: str, label: str = "保存先サブフォルダ") -> str:
    """保存先サブフォルダの安全性を判定する（単一フォルダ名または相対パス）。"""
    if not value:
        return f"{label}を入力してください"
    if re.match(r"^[A-Za-z]:", value) or value.startswith(("/", "\\")):
        return f"{label}に絶対パスは指定できません（共通保存先からの相対パスで指定してください）"
    for char in WINDOWS_FORBIDDEN_CHARS:
        if char in value:
            return f"{label}に使用できない文字 {char} が含まれています"
    segments = [seg for seg in re.split(r"[\\/]", value)]
    if any(seg == "" for seg in segments):
        return f"{label}に空のフォルダ名が含まれています"
    for segment in segments:
        error = _segment_error(segment, label)
        if error:
            return error
    return ""


def normalize_rule(raw, index: int, errors: list) -> dict:
    """1件のルールを検証して正規化する。"""
    label = f"ルール{index + 1}"
    if not isinstance(raw, dict):
        errors.append(f"{label} の形式が不正です（オブジェクトで指定してください）")
        return dict(DEFAULT_RULE)

    rule = dict(DEFAULT_RULE)
    rule["enabled"] = _as_bool(raw.get("enabled", True), f"{label} の 有効", errors)
    rule["priority"] = _as_int(raw.get("priority", index + 1), f"{label} の 優先順位", errors, 0, 99999)

    name = _as_str(raw.get("name", ""), f"{label} の ルール名", errors)
    if not name:
        errors.append(f"{label}: ルール名を入力してください")
    rule["name"] = name

    extension = _as_str(raw.get("extension", TARGET_EXTENSION), f"{label} の 拡張子", errors)
    if extension:
        if not extension.startswith("."):
            extension = "." + extension
        extension = extension.lower()
        if extension != TARGET_EXTENSION:
            errors.append(
                f"{label}: 初期仕様の対象は {TARGET_EXTENSION} のみです"
                f"（{extension} を指定したルールは一致しません）"
            )
    rule["extension"] = extension

    rule["contains_all"] = _as_str_list(raw.get("contains_all"), f"{label} の すべて含むキーワード", errors)
    rule["contains_any"] = _as_str_list(raw.get("contains_any"), f"{label} の いずれか含むキーワード", errors)
    rule["not_contains"] = _as_str_list(raw.get("not_contains"), f"{label} の 含んではいけないキーワード", errors)
    if not (rule["contains_all"] or rule["contains_any"]):
        errors.append(f"{label}: すべて含む／いずれか含む のどちらかにキーワードが必要です（全ファイルに一致してしまうため）")

    base = _as_str(raw.get("destination_base", ""), f"{label} の 保存先フォルダ", errors)
    error = base_format_error(base, f"{label} の 保存先フォルダ")
    if error:
        errors.append(error)
    rule["destination_base"] = base

    subfolder = _as_str(raw.get("destination_subfolder", ""), f"{label} の 保存先サブフォルダ", errors)
    error = subfolder_error(subfolder, f"{label} の 保存先サブフォルダ")
    if error:
        errors.append(error)
    rule["destination_subfolder"] = subfolder
    return rule


def normalize_settings(raw) -> tuple:
    """rules.json 全体を検証して正規化する。戻り値は (settings, errors)。"""
    errors: list = []
    if not isinstance(raw, dict):
        return dict(DEFAULT_SETTINGS), ["rules.json の形式が不正です（オブジェクトで指定してください）"]

    settings = dict(DEFAULT_SETTINGS)
    settings["common_base"] = _as_str(raw.get("common_base", DEFAULT_SETTINGS["common_base"]), "共通保存先フォルダ", errors)

    for field, label in (("unknown_folder", "判定不能フォルダ名"), ("log_folder", "ログフォルダ名")):
        value = _as_str(raw.get(field, DEFAULT_SETTINGS[field]), label, errors)
        error = folder_name_error(value, label)
        if error:
            errors.append(error)
        settings[field] = value

    prefix = _as_str(raw.get("evacuation_folder_prefix", DEFAULT_SETTINGS["evacuation_folder_prefix"]), "避難用フォルダ接頭語", errors)
    error = folder_name_error(prefix, "避難用フォルダ接頭語")
    if error:
        errors.append(error)
    settings["evacuation_folder_prefix"] = prefix

    for field in ("use_month_folder", "verify_after_move", "check_file_stable", "backup_rules_on_save"):
        settings[field] = _as_bool(raw.get(field, DEFAULT_SETTINGS[field]), field, errors)

    # 上書き禁止は事故防止の必須要件（仕様書 26）のため設定値によらず true に固定する。
    settings["forbid_overwrite"] = True

    for field, allowed in ENUMS.items():
        value = _as_str(raw.get(field, DEFAULT_SETTINGS[field]), field, errors)
        if value not in allowed:
            errors.append(f"{field} には {' / '.join(allowed)} のいずれかを指定してください（現在の値: {value!r}）")
            value = DEFAULT_SETTINGS[field]
        settings[field] = value

    settings["min_file_age_seconds"] = _as_int(
        raw.get("min_file_age_seconds", DEFAULT_SETTINGS["min_file_age_seconds"]), "最小ファイル経過秒数", errors, 0, 3600)
    settings["max_path_length"] = _as_int(
        raw.get("max_path_length", DEFAULT_SETTINGS["max_path_length"]), "最大パス長", errors, 60, 32767)

    raw_rules = raw.get("rules", [])
    if not isinstance(raw_rules, list):
        errors.append("rules は配列で指定してください")
        raw_rules = []
    settings["rules"] = [normalize_rule(entry, index, errors) for index, entry in enumerate(raw_rules)]
    return settings, errors


def base_format_error(text: str, label: str) -> str:
    """保存先の基準フォルダとして書式が正しいかを判定する（存在確認はしない）。"""
    if not text:
        return ""
    for char in WINDOWS_FORBIDDEN_CHARS.replace(":", ""):
        if char in text:
            return f"{label}に使用できない文字 {char} が含まれています"
    if ".." in re.split(r"[\\/]", text):
        return f"{label}に .. は使用できません"
    if not (re.match(r"^[A-Za-z]:[\\/]", text) or text.startswith(("\\\\", "/"))):
        return f"{label}は絶対パスで指定してください（例: D:\\共有\\設備記録）"
    return ""


def folder_path_error(text: str, label: str) -> str:
    """保存先の基準フォルダが実際に使えるかを判定する。使えない理由を返す。"""
    text = (text or "").strip()
    if not text or text == UNSET_BASE_PLACEHOLDER:
        return f"{label}が未設定です"
    error = base_format_error(text, label)
    if error:
        return error
    path = Path(text)
    if not path.exists():
        return f"{label}が見つかりません（{text}）"
    if not path.is_dir():
        return f"{label}がフォルダではありません（{text}）"
    return ""


def base_folder_error(settings: dict) -> str:
    """共通保存先フォルダが実際に使えるかを判定する。"""
    return folder_path_error(settings.get("common_base", ""), "共通保存先フォルダ")


def resolve_rule_base(rule: dict, settings: dict) -> tuple:
    """ルールが使う保存先フォルダと、それがルール専用かどうかを返す。"""
    own = (rule.get("destination_base") or "").strip()
    if own:
        return own, True
    return (settings.get("common_base", "") or "").strip(), False


def common_base_required(settings: dict) -> bool:
    """共通保存先フォルダを使う有効なルールがあるか（無ければ未設定でも実行できる）。"""
    return any(not (rule.get("destination_base") or "").strip() for rule in sorted_rules(settings))


# =========================================================================
# rules.json の読み書き
# =========================================================================

def load_settings(create_if_missing: bool = True) -> tuple:
    """rules.json を読み込む。戻り値は (settings, errors, fatal)。"""
    path = rules_path()
    if not path.exists():
        if not create_if_missing:
            return dict(DEFAULT_SETTINGS), [f"{RULES_FILE_NAME} が見つかりません"], True
        try:
            write_settings_atomically(dict(DEFAULT_SETTINGS), backup=False)
        except OSError as exc:
            return dict(DEFAULT_SETTINGS), [f"{RULES_FILE_NAME} を作成できません: {exc}"], True
        return dict(DEFAULT_SETTINGS), [], False

    try:
        raw = json.loads(path.read_text(encoding="utf-8"))
    except json.JSONDecodeError as exc:
        return dict(DEFAULT_SETTINGS), [f"{RULES_FILE_NAME} のJSONが不正です（{exc.lineno}行目付近: {exc.msg}）"], True
    except OSError as exc:
        return dict(DEFAULT_SETTINGS), [f"{RULES_FILE_NAME} を読み込めません: {exc}"], True

    settings, errors = normalize_settings(raw)
    return settings, errors, False


def backup_rules() -> str:
    """現在の rules.json をバックアップする。作成したパスを返す。"""
    source = rules_path()
    if not source.exists():
        return ""
    folder = app_dir() / RULES_BACKUP_FOLDER
    folder.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    target = folder / f"rules_{stamp}.json"
    index = 1
    while target.exists():
        target = folder / f"rules_{stamp}_{index:03d}.json"
        index += 1
    shutil.copy2(source, target)

    backups = sorted(folder.glob("rules_*.json"))
    for old in backups[:-RULES_BACKUP_KEEP]:
        try:
            old.unlink()
        except OSError:
            pass
    return str(target)


def write_settings_atomically(settings: dict, backup: bool = True) -> str:
    """一時ファイル経由で rules.json を安全に置換する（仕様書 13.2）。"""
    backup_path = backup_rules() if backup else ""
    target = rules_path()
    temp = target.with_suffix(".json.tmp")
    payload = json.dumps(settings, ensure_ascii=False, indent=2) + "\n"
    with open(temp, "w", encoding="utf-8", newline="\n") as handle:
        handle.write(payload)
        handle.flush()
        os.fsync(handle.fileno())
    os.replace(temp, target)
    return backup_path


# =========================================================================
# 年月抽出（仕様書 8）
# =========================================================================

# YYYY.M / YYYY.MM / YYYY-M / YYYY-MM / YYYY_M / YYYY_MM
MONTH_SEPARATED = re.compile(r"(?<!\d)((?:19|20)\d{2})([._-])(1[0-2]|0?[1-9])(?!\d)")
# YYYYMM
MONTH_COMPACT = re.compile(r"(?<!\d)((?:19|20)\d{2})(0[1-9]|1[0-2])(?!\d)")


def extract_months(file_name: str) -> list:
    """ファイル名から年月候補を抽出する。戻り値は [{"value": "2026-06", "text": "2026.06"}]。"""
    stem = Path(file_name).stem
    found = []
    spans = []

    for match in MONTH_SEPARATED.finditer(stem):
        year, month = int(match.group(1)), int(match.group(3))
        found.append({"value": f"{year:04d}-{month:02d}", "text": match.group(0)})
        spans.append(match.span())

    for match in MONTH_COMPACT.finditer(stem):
        if any(start <= match.start() < end for start, end in spans):
            continue
        year, month = int(match.group(1)), int(match.group(2))
        found.append({"value": f"{year:04d}-{month:02d}", "text": match.group(0)})

    unique = []
    seen = set()
    for entry in found:
        if entry["value"] in seen:
            continue
        seen.add(entry["value"])
        unique.append(entry)
    return unique


def format_month(value: str, month_format: str) -> str:
    """内部表現 YYYY-MM を設定の形式へ変換する。"""
    year, month = value.split("-")
    if month_format == "YYYYMM":
        return f"{year}{month}"
    if month_format == "YYYY_MM":
        return f"{year}_{month}"
    return f"{year}-{month}"


# =========================================================================
# ルール判定（仕様書 7）
# =========================================================================

def sorted_rules(settings: dict) -> list:
    """有効なルールを優先順位の昇順で返す（同値は記載順）。"""
    enabled = [(rule.get("priority", 0), index, rule)
               for index, rule in enumerate(settings.get("rules", []))
               if rule.get("enabled")]
    enabled.sort(key=lambda entry: (entry[0], entry[1]))
    return [entry[2] for entry in enabled]


def match_rule(file_name: str, rules: list) -> tuple:
    """最初に一致したルールと、その根拠を返す。戻り値は (rule, reasons)。"""
    lowered = file_name.lower()
    for rule in rules:
        extension = rule.get("extension", "")
        if extension and not lowered.endswith(extension.lower()):
            continue

        contains_all = rule.get("contains_all", [])
        if any(keyword.lower() not in lowered for keyword in contains_all):
            continue

        contains_any = rule.get("contains_any", [])
        hit_any = [keyword for keyword in contains_any if keyword.lower() in lowered]
        if contains_any and not hit_any:
            continue

        not_contains = rule.get("not_contains", [])
        if any(keyword.lower() in lowered for keyword in not_contains):
            continue

        reasons = [f"優先順位 {rule.get('priority')} のルール「{rule.get('name')}」に一致"]
        if contains_all:
            reasons.append("すべて含む: " + "、".join(contains_all))
        if hit_any:
            reasons.append("いずれか含む: " + "、".join(hit_any))
        if not_contains:
            reasons.append("含まない条件: " + "、".join(not_contains) + " を含まないことを確認")
        if extension:
            reasons.append(f"拡張子: {extension}")
        return rule, reasons
    return None, ["どのルールにも一致しませんでした（大文字小文字は区別しません）"]


# =========================================================================
# 対象ファイルの収集（仕様書 7.1 / 7.2）
# =========================================================================

def collect_target_files() -> list:
    """スクリプトと同じフォルダ内のPDFを、処理開始時点で固定して返す。"""
    files = []
    for path in sorted(app_dir().iterdir(), key=lambda item: item.name.lower()):
        if not path.is_file():
            continue
        name = path.name
        if name.lower().endswith(EXCLUDE_SUFFIXES):
            continue
        if name.startswith(EXCLUDE_PREFIXES):
            continue
        if path.suffix.lower() != TARGET_EXTENSION:
            continue
        files.append(path)
    return files


def stability_error(path: Path, settings: dict) -> str:
    """処理してよい状態かを確認する（仕様書 11）。使えない理由を返す。"""
    if not settings.get("check_file_stable"):
        return ""
    min_age = settings.get("min_file_age_seconds", 0)
    try:
        stat = path.stat()
    except OSError as exc:
        return f"ファイル情報を取得できません: {exc}"

    age = time.time() - stat.st_mtime
    if age < min_age:
        return f"更新から {int(age)} 秒しか経過していません（{min_age} 秒以上必要）"
    if stat.st_size == 0:
        return "ファイルサイズが 0 バイトです"
    try:
        with open(path, "rb+"):
            pass
    except OSError:
        return "他のアプリケーションが使用中の可能性があります"
    return ""


# =========================================================================
# 移動計画（プレビューと実行で共有する）
# =========================================================================

def evacuation_folder_name(settings: dict, today: datetime) -> str:
    return f"{settings.get('evacuation_folder_prefix', '避難用')}_{today.strftime('%Y%m%d')}"


def numbered_name(folder: Path, file_name: str) -> Path:
    """避難先で同名がある場合のみ連番を付ける（仕様書 9.4）。"""
    candidate = folder / file_name
    if not candidate.exists():
        return candidate
    stem, suffix = Path(file_name).stem, Path(file_name).suffix
    index = 1
    while index < 1000:
        candidate = folder / f"{stem}_{index:03d}{suffix}"
        if not candidate.exists():
            return candidate
        index += 1
    return folder / f"{stem}_{datetime.now().strftime('%H%M%S')}{suffix}"


def build_plan(path: Path, settings: dict, rules: list, today: datetime) -> dict:
    """1ファイルの移動計画を作る。実ファイルは変更しない。"""
    plan = {
        "file_name": path.name,
        "source": str(path),
        "extracted_month": "",
        "rule_name": "",
        "planned_destination": "",
        "destination": "",
        "actual_destination": "",
        "result": R_MATCH,
        "message": "",
        "reason": [],
        "movable": False,
    }

    stability = stability_error(path, settings)
    if stability:
        plan.update(result=R_SKIP_FILE_NOT_STABLE, message=stability,
                    reason=["処理中・同期中のファイルを壊さないためスキップします"])
        return plan

    rule, reasons = match_rule(path.name, rules)
    plan["reason"] = list(reasons)

    if rule is None:
        # ルール未一致（仕様書 15.1）
        destination = app_dir() / settings.get("unknown_folder", "_unknown") / path.name
        plan.update(rule_name="", planned_destination=str(destination), destination=str(destination),
                    result=R_UNKNOWN, message="ルールに一致しないため判定不能フォルダへ移動します", movable=True)
        return plan

    plan["rule_name"] = rule.get("name", "")

    # 保存先の基準フォルダはルールごとに指定できる（未指定なら共通保存先）
    base_text, own_base = resolve_rule_base(rule, settings)
    label = f"ルール「{rule.get('name')}」の保存先フォルダ" if own_base else "共通保存先フォルダ"
    base_error = folder_path_error(base_text, label)
    if base_error:
        plan.update(result=E_FORBIDDEN_PATH, message=base_error,
                    reason=plan["reason"] + ["保存先を確定できないため移動先を計算できません"])
        return plan

    plan["reason"].append(
        f"保存先: {'このルール専用のフォルダ' if own_base else '共通保存先'} {base_text}")
    rule_base = Path(base_text) / rule.get("destination_subfolder", "")
    month_folder = ""

    if settings.get("use_month_folder"):
        months = extract_months(path.name)
        month_format = settings.get("month_folder_format", "YYYY-MM")
        if len(months) == 1:
            plan["extracted_month"] = months[0]["value"]
            month_folder = format_month(months[0]["value"], month_format)
            plan["reason"].append(f"年月: ファイル名の「{months[0]['text']}」から {months[0]['value']} と判定")
        elif len(months) == 0:
            fallback = settings.get("month_fallback", "unknown_month")
            if fallback == "current_month":
                value = today.strftime("%Y-%m")
                plan["extracted_month"] = value
                month_folder = format_month(value, month_format)
                plan["reason"].append(f"年月: ファイル名から取得できないため実行日の {value} を使用")
                plan["message"] = "ファイル名から年月を取得できないため実行日の年月を使用します"
            elif fallback == "error":
                plan.update(result=R_MONTH_UNKNOWN,
                            message="ファイル名から年月を取得できません（設定 month_fallback=error のため移動しません）")
                plan["reason"].append("年月: 候補なし")
                return plan
            else:
                month_folder = MONTH_UNKNOWN_FOLDER
                plan.update(result=R_MONTH_UNKNOWN, message="ファイル名から年月を取得できません")
                plan["reason"].append("年月: 候補なし。通常保存先には保存せず退避します")
        else:
            month_folder = MONTH_AMBIGUOUS_FOLDER
            candidates = "、".join(f"{entry['text']}→{entry['value']}" for entry in months)
            plan.update(result=R_MONTH_AMBIGUOUS, message="年月候補が複数あります")
            plan["reason"].append(f"年月: 候補が {len(months)} 件（{candidates}）。通常保存先には保存せず退避します")
    else:
        plan["reason"].append("年月フォルダ: 使用しない設定です")

    destination_folder = rule_base / month_folder if month_folder else rule_base
    planned = destination_folder / path.name
    plan["planned_destination"] = str(planned)
    plan["destination"] = str(planned)
    plan["movable"] = True

    # パス長チェック（仕様書 14.1）
    limit = settings.get("max_path_length", 240)
    if len(str(planned)) > limit:
        plan.update(result=E_PATH_TOO_LONG, movable=False,
                    message=f"保存先パスが {len(str(planned))} 文字で上限 {limit} 文字を超えます")
        return plan

    # 同名ファイル（仕様書 9）
    if planned.exists():
        if settings.get("duplicate_mode") == "skip":
            plan.update(result=R_SKIP, movable=False, message="保存先に同名ファイルがあるため移動しません")
            plan["reason"].append("同名: 上書きしない設定のためスキップ")
            return plan
        evacuation = destination_folder / evacuation_folder_name(settings, today)
        target = numbered_name(evacuation, path.name)
        plan.update(destination=str(target), result=R_DUPLICATE_EVACUATED,
                    message="同名ファイルが存在するため避難用フォルダへ保存します")
        plan["reason"].append(f"同名: {planned} が既に存在。上書きせず避難します")
        if len(str(target)) > limit:
            plan.update(result=E_PATH_TOO_LONG, movable=False,
                        message=f"避難先パスが {len(str(target))} 文字で上限 {limit} 文字を超えます")
    return plan


def build_plans(settings: dict, files: list) -> list:
    rules = sorted_rules(settings)
    today = datetime.now()
    return [build_plan(path, settings, rules, today) for path in files]


# =========================================================================
# 安全移動（仕様書 10）
# =========================================================================

def copy_without_overwrite(source: Path, destination: Path) -> None:
    """既存ファイルを絶対に上書きしないコピー。存在すれば FileExistsError。"""
    flags = os.O_CREAT | os.O_EXCL | os.O_WRONLY | getattr(os, "O_BINARY", 0)
    handle = os.open(str(destination), flags)
    try:
        with os.fdopen(handle, "wb") as writer, open(source, "rb") as reader:
            shutil.copyfileobj(reader, writer, 1024 * 1024)
            writer.flush()
            os.fsync(writer.fileno())
    except BaseException:
        try:
            os.remove(destination)  # 書きかけを残さない
        except OSError:
            pass
        raise
    shutil.copystat(source, destination)


def safe_move(source: Path, destination: Path, settings: dict) -> tuple:
    """設定された方式でファイルを移動する。戻り値は (result_code, message, moved)。"""
    strategy = settings.get("move_strategy", "copy_verify_delete")
    verify = settings.get("verify_after_move", True)

    try:
        destination.parent.mkdir(parents=True, exist_ok=True)
    except OSError as exc:
        return E_CREATE_FOLDER, f"保存先フォルダを作成できません: {exc}", False

    try:
        source_size = source.stat().st_size
    except OSError as exc:
        return E_MOVE_FAILED, f"移動元のファイル情報を取得できません: {exc}", False

    try:
        copy_without_overwrite(source, destination)
    except FileExistsError:
        return E_MOVE_FAILED, "保存先に同名ファイルが作成されていたため中止しました（上書きはしません）", False
    except OSError as exc:
        return E_COPY_FAILED, f"コピーに失敗しました: {exc}", False

    if verify or strategy == "copy_verify_delete":
        if not destination.exists():
            return E_VERIFY_FAILED, "移動後に保存先が存在しません（移動元: 有り / 保存先: 無し）", False
        copied_size = destination.stat().st_size
        if copied_size != source_size:
            return (E_VERIFY_FAILED,
                    f"サイズが一致しません（移動元 {source_size:,} バイト / 保存先 {copied_size:,} バイト）", False)

    if strategy == "copy_only":
        return R_COPIED, "コピーのみ実行しました（移動元は残っています）", True

    try:
        os.remove(source)
    except OSError as exc:
        return E_DELETE_SOURCE_FAILED, f"コピーは成功しましたが移動元を削除できません: {exc}", True

    if strategy == "move":
        return R_MOVED, "", True
    return R_COPIED_AND_DELETED, "", True


# =========================================================================
# ログ（仕様書 16）
# =========================================================================

LOG_HEADER = ["日時", "ファイル名", "抽出年月", "判定ルール", "移動元", "通常保存予定先", "実際の保存先", "結果", "メッセージ"]


class MoveLogger:
    """UTF-8 BOM付きCSVへ追記する。"""

    def __init__(self, settings: dict):
        self.path = log_path(settings)
        self.available = True
        self.error = ""
        try:
            self.path.parent.mkdir(parents=True, exist_ok=True)
            is_new = not self.path.exists() or self.path.stat().st_size == 0
            self.handle = open(self.path, "a", encoding="utf-8-sig", newline="")
            self.writer = csv.writer(self.handle)
            if is_new:
                self.writer.writerow(LOG_HEADER)
        except OSError as exc:
            self.available = False
            self.error = f"ログを書き込めません: {exc}"

    def write(self, item: dict) -> None:
        if not self.available:
            return
        self.writer.writerow([
            now_stamp(),
            item.get("file_name", ""),
            item.get("extracted_month", ""),
            item.get("rule_name", ""),
            item.get("source", ""),
            item.get("planned_destination", ""),
            item.get("actual_destination", ""),
            item.get("result", ""),
            item.get("message", ""),
        ])
        self.handle.flush()

    def close(self) -> None:
        if self.available:
            try:
                self.handle.close()
            except OSError:
                pass


# =========================================================================
# 集計（仕様書 15.2）
# =========================================================================

SUCCESS_CODES = {R_MOVED, R_COPIED_AND_DELETED, R_COPIED}
CATEGORY_ORDER = ["success", "evacuated", "month", "unknown", "error", "skip"]
CATEGORY_LABELS = {
    "success": "移動成功",
    "evacuated": "避難保存",
    "month": "年月退避",
    "unknown": "判定不能",
    "error": "エラー",
    "skip": "スキップ",
}


def categorize(result: str) -> str:
    if result in SUCCESS_CODES or result in (R_MATCH, R_UNDONE):
        return "success"
    if result in (R_DUPLICATE_EVACUATED, R_DUPLICATE_WILL_EVACUATE):
        return "evacuated"
    if result in (R_MONTH_UNKNOWN, R_MONTH_AMBIGUOUS):
        return "month"
    if result == R_UNKNOWN:
        return "unknown"
    if result.startswith("ERROR_"):
        return "error"
    return "skip"


def summarize(items: list) -> dict:
    summary = {key: 0 for key in CATEGORY_ORDER}
    for item in items:
        summary[categorize(item.get("result", ""))] += 1
    summary["total"] = len(items)
    return summary


def summary_text(summary: dict) -> str:
    parts = [f"{CATEGORY_LABELS[key]} {summary.get(key, 0)} 件" for key in CATEGORY_ORDER]
    return " / ".join(parts)


# =========================================================================
# ロックファイル（仕様書 12）
# =========================================================================

class SortLock:
    """二重起動を防ぐロック。with 文で使う。"""

    def __init__(self):
        self.path = lock_path()
        self.acquired = False
        self.holder = ""

    def acquire(self) -> bool:
        try:
            handle = os.open(str(self.path), os.O_CREAT | os.O_EXCL | os.O_WRONLY)
        except FileExistsError:
            try:
                self.holder = self.path.read_text(encoding="utf-8").strip()
            except OSError:
                self.holder = ""
            return False
        except OSError as exc:
            self.holder = f"ロックファイルを作成できません: {exc}"
            return False
        with os.fdopen(handle, "w", encoding="utf-8") as writer:
            writer.write(f"pid={os.getpid()} started={now_stamp()}\n")
        self.acquired = True
        return True

    def release(self) -> None:
        if not self.acquired:
            return
        try:
            self.path.unlink()
        except OSError:
            pass
        self.acquired = False

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc, traceback):
        self.release()
        return False


# =========================================================================
# 仕分け実行
# =========================================================================

def run_sort(settings: dict) -> dict:
    """PDF移動を実行する。戻り値は {"ok", "items", "summary", "message", ...}。"""
    files = collect_target_files()
    plans = build_plans(settings, files)
    logger = MoveLogger(settings)
    today = datetime.now()
    items = []
    undo_records = []

    for plan in plans:
        item = dict(plan)
        source = Path(plan["source"])

        if not plan["movable"]:
            item["actual_destination"] = ""
            logger.write(item)
            items.append(item)
            continue

        destination = Path(plan["destination"])

        # 計画作成後に状況が変わっている可能性があるため、直前に再確認する。
        if plan["result"] != R_DUPLICATE_EVACUATED and destination.exists():
            if settings.get("duplicate_mode") == "skip":
                item.update(result=R_SKIP, actual_destination="",
                            message="保存先に同名ファイルが作成されたため移動しませんでした")
                logger.write(item)
                items.append(item)
                continue
            evacuation = destination.parent / evacuation_folder_name(settings, today)
            destination = numbered_name(evacuation, source.name)
            item.update(result=R_DUPLICATE_EVACUATED, destination=str(destination),
                        message="同名ファイルが存在したため避難保存")

        result, message, moved = safe_move(source, destination, settings)

        if result in SUCCESS_CODES:
            # 分類を保ったまま、実行方式に応じた結果コードへ置き換える。
            if item["result"] in (R_UNKNOWN, R_MONTH_UNKNOWN, R_MONTH_AMBIGUOUS, R_DUPLICATE_EVACUATED):
                pass  # 分類コードを優先して残す（仕様書 17）
            else:
                item["result"] = result
            item["actual_destination"] = str(destination)
            if message:
                item["message"] = (item["message"] + " / " if item["message"] else "") + message
            if moved and result != R_COPIED:
                undo_records.append({"source": str(source), "destination": str(destination),
                                     "file_name": source.name})
        else:
            item["result"] = result
            item["actual_destination"] = str(destination) if moved else ""
            item["message"] = message
            if moved:
                undo_records.append({"source": str(source), "destination": str(destination),
                                     "file_name": source.name})

        logger.write(item)
        items.append(item)

    logger.close()
    summary = summarize(items)
    run_info = {
        "started_at": now_stamp(),
        "records": undo_records,
        "undone": False,
    }
    if undo_records:
        save_last_run(settings, run_info)

    return {
        "ok": True,
        "items": items,
        "summary": summary,
        "log_path": str(log_path(settings)),
        "log_error": logger.error,
        "undo_count": len(undo_records),
        "started_at": run_info["started_at"],
    }


def save_last_run(settings: dict, run_info: dict) -> None:
    try:
        path = last_run_path(settings)
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(json.dumps(run_info, ensure_ascii=False, indent=2), encoding="utf-8")
    except OSError:
        pass


def load_last_run(settings: dict) -> dict:
    path = last_run_path(settings)
    if not path.exists():
        return {}
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return {}
    return data if isinstance(data, dict) else {}


def run_undo(settings: dict) -> dict:
    """直前の実行で移動したファイルを元の場所へ戻す（仕様書 33 / 設計原則 5）。"""
    run_info = load_last_run(settings)
    records = run_info.get("records") or []
    if not records or run_info.get("undone"):
        return {"ok": False, "message": "取り消せる移動がありません", "items": [], "summary": summarize([])}

    logger = MoveLogger(settings)
    items = []
    restored = 0

    for record in reversed(records):
        moved_to = Path(record["destination"])
        original = Path(record["source"])
        item = {
            "file_name": record.get("file_name", moved_to.name),
            "extracted_month": "",
            "rule_name": "",
            "source": str(moved_to),
            "planned_destination": str(original),
            "actual_destination": "",
            "result": R_UNDONE,
            "message": "",
            "reason": [],
        }
        if not moved_to.exists():
            item.update(result=E_UNDO_FAILED, message="移動先にファイルがありません（既に移動された可能性があります）")
        elif original.exists():
            item.update(result=E_UNDO_FAILED, message="元の場所に同名ファイルがあるため戻せません（上書きはしません）")
        else:
            result, message, _ = safe_move(moved_to, original, settings)
            if result in SUCCESS_CODES:
                item.update(result=R_UNDONE, actual_destination=str(original), message="移動を取り消しました")
                restored += 1
            else:
                item.update(result=E_UNDO_FAILED, message=message)
        logger.write(item)
        items.append(item)

    logger.close()
    run_info["undone"] = True
    save_last_run(settings, run_info)
    return {
        "ok": restored > 0,
        "items": items,
        "restored": restored,
        "failed": len(items) - restored,
        "message": f"{restored} 件を元の場所へ戻しました",
        "log_path": str(log_path(settings)),
    }


# =========================================================================
# 画面へ渡す付帯情報
# =========================================================================

MONTH_PATTERN_EXAMPLES = ["2026.6", "2026.06", "2026-6", "2026-06", "2026_6", "2026_06", "202606"]


def build_meta(settings: dict, errors: list) -> dict:
    """画面表示用の付帯情報（出どころ・件数・実行可否の根拠）をまとめる。"""
    files = []
    for path in collect_target_files():
        try:
            stat = path.stat()
            size, modified = stat.st_size, datetime.fromtimestamp(stat.st_mtime).strftime("%Y-%m-%d %H:%M")
        except OSError:
            size, modified = 0, ""
        files.append({
            "name": path.name,
            "size": size,
            "modified": modified,
            "blocked": stability_error(path, settings),
        })

    lock = lock_path()
    holder = ""
    if lock.exists():
        try:
            holder = lock.read_text(encoding="utf-8").strip()
        except OSError:
            holder = ""

    run_info = load_last_run(settings)
    undo_available = bool(run_info.get("records")) and not run_info.get("undone")

    return {
        "app_version": APP_VERSION,
        "script_folder": str(app_dir()),
        "rules_path": str(rules_path()),
        "log_path": str(log_path(settings)),
        "unknown_path": str(app_dir() / settings.get("unknown_folder", "_unknown")),
        "backup_folder": str(app_dir() / RULES_BACKUP_FOLDER),
        "target_extension": TARGET_EXTENSION,
        "files": files,
        "file_count": len(files),
        "base_folder_error": base_folder_error(settings),
        "common_base_required": common_base_required(settings),
        "errors": errors,
        "lock": {"locked": lock.exists(), "holder": holder, "path": str(lock)},
        "undo": {
            "available": undo_available,
            "count": len(run_info.get("records") or []),
            "started_at": run_info.get("started_at", ""),
        },
        "month_patterns": MONTH_PATTERN_EXAMPLES,
        "month_unknown_folder": MONTH_UNKNOWN_FOLDER,
        "month_ambiguous_folder": MONTH_AMBIGUOUS_FOLDER,
        "evacuation_example": f"{settings.get('evacuation_folder_prefix', '避難用')}_{datetime.now().strftime('%Y%m%d')}",
        "generated_at": now_stamp(),
    }


# =========================================================================
# コンソール表示
# =========================================================================

CONSOLE_LABELS = {
    R_MOVED: "[ OK ] 移動成功",
    R_COPIED_AND_DELETED: "[ OK ] 移動成功",
    R_COPIED: "[ OK ] コピー",
    R_MATCH: "[ OK ] 移動予定",
    R_DUPLICATE_EVACUATED: "[ ! ] 避難保存",
    R_DUPLICATE_WILL_EVACUATE: "[ ! ] 避難予定",
    R_UNKNOWN: "[ ? ] 判定不能",
    R_MONTH_UNKNOWN: "[ ! ] 年月不明",
    R_MONTH_AMBIGUOUS: "[ ! ] 年月確認要",
    R_SKIP: "[ - ] スキップ",
    R_SKIP_FILE_NOT_STABLE: "[ - ] スキップ",
    R_UNDONE: "[ OK ] 取り消し",
}


def console_label(result: str) -> str:
    return CONSOLE_LABELS.get(result, "[ NG ] エラー")


def print_items(items: list) -> None:
    if not items:
        echo("  対象のPDFはありません。")
        return
    for item in items:
        destination = item.get("actual_destination") or item.get("destination") or "（移動しません）"
        echo(f"  {console_label(item['result'])}  {item['file_name']}")
        echo(f"      結果コード : {item['result']}")
        echo(f"      判定ルール : {item['rule_name'] or '（一致なし）'}")
        echo(f"      抽出年月   : {item['extracted_month'] or '（取得できません）'}")
        echo(f"      保存先     : {destination}")
        if item.get("message"):
            echo(f"      メッセージ : {item['message']}")
        for reason in item.get("reason", []):
            echo(f"      根拠       : {reason}")
        echo()


def print_header(title: str) -> None:
    echo("=" * 72)
    echo(f" {title}")
    echo("=" * 72)


def print_settings_errors(errors: list) -> None:
    echo(f"[ NG ] {RULES_FILE_NAME} に問題があります（{len(errors)} 件）")
    for error in errors:
        echo(f"      - {error}")


def load_for_cli() -> tuple:
    """CLI用に rules.json を読み込む。問題があれば理由を表示して None を返す。"""
    settings, errors, fatal = load_settings()
    if fatal:
        print_settings_errors(errors)
        echo()
        echo(f"      対象ファイル: {rules_path()}")
        echo("      修正するか、ファイルを削除してから再実行してください（削除時は初期設定で自動作成されます）。")
        return None, errors
    if errors:
        print_settings_errors(errors)
        echo()
        echo("      ルール編集画面（edit_rules.bat）で修正してください。処理は中断します。")
        return None, errors
    return settings, []


# =========================================================================
# CLI: preview / sort
# =========================================================================

def cli_preview() -> int:
    print_header(f"{APP_NAME} - 移動プレビュー（画面を使わない実行）")
    settings, _ = load_for_cli()
    if settings is None:
        return 1

    if common_base_required(settings):
        base_error = base_folder_error(settings)
        if base_error:
            echo(f"[ NG ] {base_error}")
            echo(f"      共通保存先フォルダ: {settings.get('common_base')}")
            echo("      PDF仕分けツール.bat を実行して共通設定を修正してください。")
            echo("      （ルールごとに保存先フォルダを指定している場合、共通保存先は使われません）")
            return 1

    files = collect_target_files()
    echo(f"対象フォルダ : {app_dir()}")
    echo(f"対象PDF      : {len(files)} 件（サブフォルダ内は対象外）")
    echo(f"共通保存先   : {settings['common_base']}")
    echo()
    items = build_plans(settings, files)
    print_items(items)
    echo(summary_text(summarize(items)))
    echo("※ プレビューのためファイルは移動していません。")
    return 0


def cli_sort() -> int:
    print_header(f"{APP_NAME} - PDF移動（画面を使わない実行）")
    settings, _ = load_for_cli()
    if settings is None:
        return 1

    if common_base_required(settings):
        base_error = base_folder_error(settings)
        if base_error:
            echo(f"[ NG ] {base_error}")
            echo(f"      共通保存先フォルダ: {settings.get('common_base')}")
            echo("      PDF仕分けツール.bat を実行して共通設定を修正してください。")
            echo("      （ルールごとに保存先フォルダを指定している場合、共通保存先は使われません）")
            return 1

    with SortLock() as lock:
        if not lock.acquire():
            echo(f"[ NG ] {E_LOCKED}: 他の処理が実行中の可能性があります。")
            echo(f"      ロックファイル: {lock_path()}")
            if lock.holder:
                echo(f"      作成情報      : {lock.holder}")
            echo("      実行中のウィンドウが無い場合は、上記ファイルを削除してから再実行してください。")
            return 1

        echo(f"対象フォルダ : {app_dir()}")
        echo(f"共通保存先   : {settings['common_base']}")
        echo()
        result = run_sort(settings)

    print_items(result["items"])
    echo("-" * 72)
    echo(summary_text(result["summary"]))
    echo(f"ログ         : {result['log_path']}")
    if result.get("log_error"):
        echo(f"[ NG ] {result['log_error']}")
    if result["summary"]["total"] == 0:
        echo("対象のPDFがありません。仕分けしたいPDFをこのフォルダへ置いてから実行してください。")
    return 0


# =========================================================================
# ローカルWebサーバー（仕様書 18 / 19）
# =========================================================================

SORT_GUARD = threading.Lock()
MAX_BODY_BYTES = 5 * 1024 * 1024


class SorterHandler(BaseHTTPRequestHandler):
    server_version = f"PdfSorter/{APP_VERSION}"
    protocol_version = "HTTP/1.1"

    # --- 応答ヘルパー ---------------------------------------------------
    def log_message(self, format_string, *args):  # noqa: A002 - 標準ログは抑止する
        return

    def _allowed_hosts(self) -> set:
        port = self.server.server_address[1]
        return {f"127.0.0.1:{port}", f"localhost:{port}", f"[::1]:{port}"}

    def _local_only(self) -> bool:
        """ローカル以外からのアクセスを拒否する（DNSリバインディング対策）。"""
        if self.client_address[0] not in ("127.0.0.1", "::1"):
            return False
        if self.headers.get("Host", "") not in self._allowed_hosts():
            return False
        origin = self.headers.get("Origin")
        if origin and origin.split("//")[-1] not in self._allowed_hosts():
            return False
        return True

    def _send_bytes(self, status: int, body: bytes, content_type: str) -> None:
        self.send_response(status)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        self.send_header("X-Content-Type-Options", "nosniff")
        self.send_header("Content-Security-Policy",
                         "default-src 'self'; img-src 'self' data:; base-uri 'none'; form-action 'none'")
        self.send_header("Referrer-Policy", "no-referrer")
        self.end_headers()
        if self.command != "HEAD":
            self.wfile.write(body)

    def _send_text(self, status: int, text: str, content_type: str) -> None:
        self._send_bytes(status, text.encode("utf-8"), f"{content_type}; charset=utf-8")

    def _send_json(self, payload: dict, status: int = 200) -> None:
        self._send_bytes(status, json.dumps(payload, ensure_ascii=False).encode("utf-8"),
                         "application/json; charset=utf-8")

    def _read_json(self):
        try:
            length = int(self.headers.get("Content-Length", "0"))
        except ValueError:
            return None, "Content-Length が不正です"
        if length <= 0:
            return None, "リクエストボディがありません"
        if length > MAX_BODY_BYTES:
            return None, "リクエストボディが大きすぎます"
        try:
            raw = self.rfile.read(length).decode("utf-8")
        except (OSError, UnicodeDecodeError) as exc:
            return None, f"リクエストを読み取れません: {exc}"
        try:
            return json.loads(raw), ""
        except json.JSONDecodeError as exc:
            return None, f"JSONを解釈できません: {exc.msg}"

    # --- ルーティング ---------------------------------------------------
    def do_GET(self):  # noqa: N802 - BaseHTTPRequestHandler の規約
        if not self._local_only():
            self._send_text(403, "Forbidden", "text/plain")
            return
        path = self.path.split("?")[0]
        if path == "/":
            self._send_text(200, UI_HTML, "text/html")
        elif path == "/app.css":
            self._send_text(200, UI_CSS, "text/css")
        elif path == "/app.js":
            self._send_text(200, UI_JS, "text/javascript")
        elif path in ("/api/rules", "/api/status"):
            settings, errors, fatal = load_settings()
            payload = dict(settings)
            payload["_meta"] = build_meta(settings, errors)
            payload["_errors"] = errors
            payload["_fatal"] = fatal
            self._send_json(payload)
        elif path == "/favicon.ico":
            self._send_bytes(204, b"", "image/x-icon")
        else:
            self._send_json({"ok": False, "message": "見つかりません"}, 404)

    def do_HEAD(self):  # noqa: N802
        self.do_GET()

    def do_POST(self):  # noqa: N802
        if not self._local_only():
            self._send_text(403, "Forbidden", "text/plain")
            return
        path = self.path.split("?")[0]
        if path == "/api/rules":
            self._handle_save()
        elif path == "/api/preview":
            self._handle_preview()
        elif path == "/api/sort":
            self._handle_sort()
        elif path == "/api/undo":
            self._handle_undo()
        elif path == "/api/shutdown":
            self._handle_shutdown()
        else:
            self._send_json({"ok": False, "message": "見つかりません"}, 404)

    # --- API 本体 -------------------------------------------------------
    def _receive_settings(self):
        """受信JSONを検証して正規化する。戻り値は (settings, error_response)。"""
        raw, error = self._read_json()
        if error:
            return None, {"ok": False, "result": E_RULES_JSON_INVALID, "message": error, "errors": [error]}
        if isinstance(raw, dict):
            raw = {key: value for key, value in raw.items() if not key.startswith("_")}
        settings, errors = normalize_settings(raw)
        if errors:
            return None, {"ok": False, "result": E_RULE_INVALID,
                          "message": f"設定に {len(errors)} 件の問題があります。保存も実行もしていません。",
                          "errors": errors}
        return settings, None

    def _handle_save(self):
        settings, error = self._receive_settings()
        if error:
            self._send_json(error, 400)
            return
        try:
            backup = write_settings_atomically(settings, backup=settings.get("backup_rules_on_save", True))
        except OSError as exc:
            self._send_json({"ok": False, "result": E_RULES_JSON_INVALID,
                             "message": f"{RULES_FILE_NAME} を保存できません: {exc}", "errors": [str(exc)]}, 500)
            return
        self._send_json({
            "ok": True,
            "message": "saved",
            "saved_at": now_stamp(),
            "backup_path": backup,
            "_meta": build_meta(settings, []),
        })

    def _handle_preview(self):
        settings, error = self._receive_settings()
        if error:
            self._send_json(error, 400)
            return
        base_error = base_folder_error(settings) if common_base_required(settings) else ""
        files = collect_target_files()
        items = build_plans(settings, files)
        self._send_json({
            "ok": not base_error,
            "items": items,
            "summary": summarize(items),
            "message": base_error or f"{len(items)} 件のPDFを判定しました（ファイルは移動していません）",
            "base_folder_error": base_error,
            "_meta": build_meta(settings, []),
        })

    def _handle_sort(self):
        settings, error = self._receive_settings()
        if error:
            self._send_json(error, 400)
            return
        if common_base_required(settings):
            base_error = base_folder_error(settings)
            if base_error:
                self._send_json({"ok": False, "result": E_FORBIDDEN_PATH, "message": base_error,
                                 "errors": [base_error]}, 400)
                return
        try:
            write_settings_atomically(settings, backup=settings.get("backup_rules_on_save", True))
        except OSError as exc:
            self._send_json({"ok": False, "result": E_RULES_JSON_INVALID,
                             "message": f"{RULES_FILE_NAME} を保存できないため実行を中止しました: {exc}",
                             "errors": [str(exc)]}, 500)
            return

        with SORT_GUARD:
            with SortLock() as lock:
                if not lock.acquire():
                    self._send_json({
                        "ok": False, "result": E_LOCKED,
                        "message": "他の処理が実行中のため開始できません",
                        "errors": [f"ロックファイル {lock_path()} が存在します。{lock.holder}"],
                    }, 409)
                    return
                result = run_sort(settings)

        result["_meta"] = build_meta(settings, [])
        result["message"] = summary_text(result["summary"])
        self._send_json(result)

    def _handle_undo(self):
        settings, errors, fatal = load_settings()
        if fatal:
            self._send_json({"ok": False, "result": E_RULES_JSON_INVALID,
                             "message": "設定を読み込めないため取り消せません", "errors": errors}, 400)
            return
        with SORT_GUARD:
            with SortLock() as lock:
                if not lock.acquire():
                    self._send_json({"ok": False, "result": E_LOCKED,
                                     "message": "他の処理が実行中のため取り消せません",
                                     "errors": [f"ロックファイル {lock_path()} が存在します。"]}, 409)
                    return
                result = run_undo(settings)
        result["_meta"] = build_meta(settings, errors)
        self._send_json(result)

    def _handle_shutdown(self):
        """画面からツールを終了する（コンソールを操作させないため）。"""
        self._send_json({"ok": True, "message": "ツールを終了しました"})
        threading.Thread(target=self.server.shutdown, daemon=True).start()


def start_server() -> int:
    print_header(APP_NAME)
    _, errors, fatal = load_settings()  # 起動時に rules.json を用意する
    if errors:
        print_settings_errors(errors)
        echo()
        if fatal:
            echo(f"      対象ファイル: {rules_path()}")
        echo("      画面は起動しますが、修正するまで保存と実行はできません。")
        echo()

    httpd = None
    for offset in range(PORT_RETRY):
        try:
            httpd = ThreadingHTTPServer((HOST, PORT + offset), SorterHandler)
            break
        except OSError:
            continue
    if httpd is None:
        echo(f"[ NG ] ポート {PORT}〜{PORT + PORT_RETRY - 1} がすべて使用中のため起動できません。")
        echo("      起動済みの操作画面を終了してから、もう一度実行してください。")
        return 1

    url = f"http://{HOST}:{httpd.server_address[1]}/"
    echo(f"対象フォルダ : {app_dir()}")
    echo(f"設定ファイル : {rules_path()}")
    echo(f"操作画面     : {url}")
    echo()
    echo("ブラウザで操作画面を開きます。")
    echo("ルールの設定、移動プレビュー、PDF移動は、すべて画面から行えます。")
    echo("終了するときは、画面右上の「…」から「ツールを終了」を選んでください。")
    echo("（このサーバーは 127.0.0.1 のみで待ち受けます。外部へは公開されません）")
    echo()

    threading.Timer(0.6, lambda: webbrowser.open(url)).start()
    try:
        httpd.serve_forever()
    except KeyboardInterrupt:
        echo()
    finally:
        httpd.server_close()
    echo("ツールを終了しました。このウィンドウは閉じて構いません。")
    return 0


# =========================================================================
# エントリポイント
# =========================================================================

USAGE = """使い方:
  py pdf_sorter_app.py           操作画面を開く（PDF仕分けツール.bat と同じ。通常はこちら）
  py pdf_sorter_app.py server    操作画面を開く
  py pdf_sorter_app.py sort      画面を使わずPDFを移動する（タスクスケジューラ等での自動実行用）
  py pdf_sorter_app.py preview   画面を使わず移動予定先をコンソールに表示する（移動しません）
"""


def main(argv: list) -> int:
    setup_console()
    mode = (argv[1] if len(argv) > 1 else "server").lower()
    if mode in ("server", "edit"):
        return start_server()
    if mode == "sort":
        return cli_sort()
    if mode == "preview":
        return cli_preview()
    if mode in ("-h", "--help", "help"):
        echo(USAGE)
        return 0
    echo(f"[ NG ] 不明なモード: {mode}")
    echo()
    echo(USAGE)
    return 1


# =========================================================================
# 画面アセット
#   将来 HTML / CSS / JavaScript を別ファイルへ分離できるよう、
#   ここから下は独立した文字列として保持している（仕様書 27.2）。
# =========================================================================

UI_HTML = r"""<!DOCTYPE html>
<html lang="ja">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>PDF自動仕分けツール</title>
<link rel="stylesheet" href="/app.css">
</head>
<body>
<div class="app">

  <header class="topbar">
    <div class="topbar__identity">
      <span class="topbar__mark" aria-hidden="true">PDF</span>
      <div class="topbar__names">
        <h1 class="topbar__title">PDF自動仕分けツール</h1>
        <p class="topbar__sub" id="topbar-sub">読み込み中です</p>
      </div>
    </div>
    <dl class="facts" id="facts"></dl>
    <div class="topbar__actions">
      <button type="button" class="btn btn--quiet" id="btn-settings" aria-haspopup="dialog">共通設定</button>
      <div class="menu">
        <button type="button" class="btn btn--icon" id="btn-menu" aria-haspopup="menu" aria-expanded="false"
                aria-label="その他の操作">···</button>
        <div class="menu__list" id="menu-list" role="menu" hidden>
          <button type="button" role="menuitem" data-menu="reload">rules.json を再読み込み</button>
          <button type="button" role="menuitem" data-menu="copy-log">ログのパスをコピー</button>
          <button type="button" role="menuitem" data-menu="copy-backup">バックアップ先のパスをコピー</button>
          <hr class="menu__split">
          <button type="button" role="menuitem" data-menu="clear-rules" class="menu__danger">すべてのルールを削除</button>
          <button type="button" role="menuitem" data-menu="quit">ツールを終了</button>
        </div>
      </div>
    </div>
  </header>

  <main class="workspace">

    <section class="panel panel--rules" aria-labelledby="rules-title">
      <div class="panel__head">
        <div class="panel__heading">
          <h2 class="panel__title" id="rules-title">仕分けルール</h2>
          <p class="panel__note" id="rules-note">—</p>
        </div>
        <button type="button" class="btn btn--quiet" id="btn-add-rule">ルールを追加</button>
      </div>
      <div class="panel__body scroll" id="rule-list"></div>
      <p class="panel__foot">上から優先順に判定し、最初に一致したルールを使います。大文字・小文字は区別しません。</p>
    </section>

    <section class="panel panel--result" aria-labelledby="result-title">
      <div class="panel__head">
        <div class="panel__heading">
          <h2 class="panel__title" id="result-title">判定結果</h2>
          <p class="panel__note" id="result-note">—</p>
        </div>
        <div class="panel__head-actions" id="result-head-actions"></div>
      </div>
      <div class="summary" id="summary" hidden></div>
      <div class="undo-bar" id="undo-bar" hidden></div>
      <div class="panel__body scroll" id="result-area"></div>
      <p class="panel__foot" id="result-foot">対象はこのフォルダ直下の .pdf のみです。サブフォルダ内と一時ファイルは対象外です。</p>
    </section>

  </main>

  <footer class="actionbar">
    <div class="actionbar__status">
      <p class="actionbar__reason" id="cta-reason">—</p>
      <p class="actionbar__hint" id="cta-hint"></p>
    </div>
    <div class="actionbar__buttons">
      <button type="button" class="btn btn--ghost" id="btn-reload">再読み込み</button>
      <button type="button" class="btn btn--quiet" id="btn-save">rules.json に保存</button>
      <button type="button" class="btn btn--quiet" id="btn-preview">移動プレビュー</button>
      <button type="button" class="btn btn--primary" id="btn-sort">読み込み中</button>
    </div>
  </footer>

</div>

<div class="toasts" id="toasts" aria-live="polite"></div>

<div class="farewell" id="farewell" hidden>
  <div class="farewell__box">
    <p class="farewell__title">ツールを終了しました</p>
    <p class="farewell__text">この画面（ブラウザのタブ）は閉じて構いません。<br>
      もう一度使うときは <b>PDF仕分けツール.bat</b> を実行してください。</p>
  </div>
</div>

<dialog class="dialog" id="settings-dialog" aria-labelledby="settings-title">
  <form method="dialog" class="dialog__form">
    <header class="dialog__head">
      <h2 id="settings-title">共通設定</h2>
      <p>すべてのルールに共通する設定です。変更は画面上の内容として保持され、保存または実行のときに rules.json へ書き込まれます。</p>
      <p class="dialog__lead" id="settings-lead" hidden>まず、仕分けしたPDFの保存先（共通保存先フォルダ）を指定してください。ここが決まると実行できるようになります。</p>
    </header>

    <div class="dialog__body scroll">
      <div class="field field--wide">
        <label for="set-common-base">共通保存先フォルダ</label>
        <input type="text" id="set-common-base" data-setting="common_base" spellcheck="false"
               placeholder="例: D:\共有\設備記録">
        <p class="field__hint" id="hint-common-base">仕分け先の親フォルダです。絶対パスで指定します。</p>
      </div>

      <div class="field">
        <label for="set-unknown-folder">判定不能フォルダ名</label>
        <input type="text" id="set-unknown-folder" data-setting="unknown_folder" spellcheck="false">
        <p class="field__hint" id="hint-unknown-folder">ルールに一致しないPDFの退避先です。</p>
      </div>

      <div class="field">
        <label for="set-log-folder">ログフォルダ名</label>
        <input type="text" id="set-log-folder" data-setting="log_folder" spellcheck="false">
        <p class="field__hint" id="hint-log-folder">移動結果CSVの保存先です。</p>
      </div>

      <div class="field field--wide">
        <label class="check">
          <input type="checkbox" id="set-use-month" data-setting="use_month_folder">
          <span class="check__box" aria-hidden="true"></span>
          <span class="check__text">年月フォルダを使用する</span>
        </label>
        <p class="field__hint" id="hint-use-month">保存先サブフォルダの下に YYYY-MM フォルダを作成します。</p>
      </div>

      <div class="field">
        <label for="set-month-fallback">年月が取れない場合の動作</label>
        <select id="set-month-fallback" data-setting="month_fallback">
          <option value="unknown_month">_年月不明 フォルダへ退避する（推奨）</option>
          <option value="current_month">実行日の年月フォルダへ保存する</option>
          <option value="error">移動せずエラーにする</option>
        </select>
        <p class="field__hint" id="hint-month-fallback"></p>
      </div>

      <div class="field">
        <label for="set-duplicate-mode">同名ファイル時の処理</label>
        <select id="set-duplicate-mode" data-setting="duplicate_mode">
          <option value="evacuate">避難用フォルダへ保存する（推奨）</option>
          <option value="skip">移動せずスキップする</option>
        </select>
        <p class="field__hint" id="hint-duplicate-mode"></p>
      </div>

      <div class="field field--wide">
        <label for="set-move-strategy">安全移動方式</label>
        <select id="set-move-strategy" data-setting="move_strategy">
          <option value="copy_verify_delete">コピー → 検証 → 元を削除（推奨）</option>
          <option value="move">移動（コピー後に元を削除。検証なし）</option>
          <option value="copy_only">コピーのみ（元ファイルを残す）</option>
        </select>
        <p class="field__hint" id="hint-move-strategy"></p>
      </div>

      <div class="field field--wide">
        <label class="check">
          <input type="checkbox" id="set-check-stable" data-setting="check_file_stable">
          <span class="check__box" aria-hidden="true"></span>
          <span class="check__text">処理前にファイル安定確認をする</span>
        </label>
        <p class="field__hint">スキャン中・同期中の未完成ファイルを処理しないための確認です。</p>
      </div>

      <div class="field">
        <label for="set-min-age">最小ファイル経過秒数</label>
        <div class="field__unit">
          <input type="number" id="set-min-age" data-setting="min_file_age_seconds" min="0" max="3600" step="1">
          <span>秒</span>
        </div>
        <p class="field__hint">更新からこの秒数が経つまで処理しません。</p>
      </div>

      <div class="field">
        <label for="set-max-path">最大パス長</label>
        <div class="field__unit">
          <input type="number" id="set-max-path" data-setting="max_path_length" min="60" max="32767" step="1">
          <span>文字</span>
        </div>
        <p class="field__hint">保存先のフルパスがこれを超えると移動しません。</p>
      </div>

      <div class="fixed-list field--wide">
        <h3>変更できない設定</h3>
        <ul id="fixed-settings"></ul>
      </div>
    </div>

    <footer class="dialog__foot">
      <p class="dialog__foot-note" id="settings-foot-note"></p>
      <button type="submit" class="btn btn--quiet" id="btn-close-settings">閉じる</button>
    </footer>
  </form>
</dialog>

<script src="/app.js"></script>
</body>
</html>
"""


UI_CSS = r""":root {
  /* 色 --------------------------------------------------------------
     メインカラー: 深い青 #123a68（上部バーと主要動作）
     ベースカラー: 淡い青みのグレー（面）
     アクセントカラー: 深いティール #0b6d7d（状態・フォーカスなど小面積）
     文字色・状態色はいずれも背景に対して 4.5:1 以上を確保している。 */
  --color-main: #123a68;
  --color-main-hover: #1a4d86;
  --color-main-ink: #ffffff;
  --color-on-main: #ffffff;
  --color-on-main-muted: #b9cbe4;
  --color-on-main-quiet: rgba(255, 255, 255, 0.10);
  --color-on-main-quiet-strong: rgba(255, 255, 255, 0.18);
  --color-on-main-line: rgba(255, 255, 255, 0.26);
  --color-on-main-warn: #ffd694;
  --color-on-main-warn-quiet: rgba(255, 214, 148, 0.16);

  --color-accent: #0b6d7d;
  --color-accent-quiet: rgba(11, 109, 125, 0.10);
  --color-accent-edge: rgba(11, 109, 125, 0.30);
  --color-focus: #123a68;

  --color-canvas: #e4ebf5;
  --color-surface: #ffffff;
  --color-surface-2: #f1f5fb;
  --color-surface-3: #e2eaf5;
  --color-field: #ffffff;
  --color-line: #c2d1e4;
  --color-line-quiet: #dbe4f0;
  --color-text: #10273f;
  --color-text-muted: #35496a;
  --color-text-faint: #4b6280;

  --color-success: #0f6a4a;
  --color-success-quiet: rgba(15, 106, 74, 0.10);
  --color-success-edge: rgba(15, 106, 74, 0.28);
  --color-warn: #8a5300;
  --color-warn-quiet: rgba(138, 83, 0, 0.10);
  --color-warn-edge: rgba(138, 83, 0, 0.28);
  --color-danger: #b02318;
  --color-danger-quiet: rgba(176, 35, 24, 0.09);
  --color-danger-edge: rgba(176, 35, 24, 0.28);
  --color-info: #0b6d7d;
  --color-info-quiet: rgba(11, 109, 125, 0.10);
  --color-info-edge: rgba(11, 109, 125, 0.30);
  --color-neutral: #47607f;
  --color-neutral-quiet: rgba(71, 96, 127, 0.09);
  --color-neutral-edge: rgba(71, 96, 127, 0.26);

  --color-bar-shade: #dde6f2;
  --color-well: rgba(18, 58, 104, 0.045);
  --color-well-strong: rgba(18, 58, 104, 0.075);
  --color-row-hover: rgba(11, 109, 125, 0.06);
  --color-backdrop: rgba(16, 39, 63, 0.45);
  --color-placeholder: #5f7797;

  /* 余白 ----------------------------------------------------------- */
  --space-hair: 1px;
  --space-0: 2px;
  --space-1: 4px;
  --space-2: 8px;
  --space-3: 12px;
  --space-4: 16px;
  --space-5: 24px;

  /* 形 ------------------------------------------------------------- */
  --radius-xs: 3px;
  --radius-sm: 6px;
  --radius-md: 10px;
  --radius-lg: 14px;
  --radius-pill: 999px;
  --border-width: 1px;
  --focus-ring: 2px;
  --focus-offset: 2px;

  /* 文字 ----------------------------------------------------------- */
  --font-base: "Segoe UI", "Yu Gothic UI", "Hiragino Sans", "Noto Sans JP", system-ui, sans-serif;
  --font-mono: "Cascadia Mono", "Consolas", "BIZ UDGothic", "SFMono-Regular", monospace;
  --font-size-2xs: 9px;
  --font-size-icon: 10px;
  --font-size-xs: 11px;
  --font-size-sm: 12px;
  --font-size-md: 13px;
  --font-size-lg: 15px;
  --line-tight: 1.35;
  --line-normal: 1.6;

  /* 大きさ --------------------------------------------------------- */
  --control-height: 30px;
  --control-height-lg: 38px;
  --topbar-height: 60px;
  --actionbar-height: 62px;
  --mark-size: 34px;
  --scrollbar-size: 10px;
  --scrollbar-inset: 3px;
  --fact-max-width: 340px;
  --menu-min-width: 240px;
  --panel-min-rules: 360px;
  --panel-min-result: 420px;
  --dialog-width: 760px;
  --input-width-number: 110px;
  --input-width-priority: 46px;
  --switch-width: 34px;
  --switch-height: 18px;
  --switch-knob: 12px;
  --switch-travel: 16px;
  --stepper-width: 20px;
  --stepper-height: 14px;
  --lift: 6px;

  /* 動き ----------------------------------------------------------- */
  --duration-fast: 110ms;
  --duration: 180ms;
  --easing: cubic-bezier(0.2, 0.7, 0.3, 1);

  /* 影 ------------------------------------------------------------- */
  --shadow-2: 0 12px 30px rgba(15, 41, 71, 0.18);
}

* { box-sizing: border-box; }

[hidden] { display: none !important; }

html, body {
  height: 100%;
  margin: 0;
  overflow: hidden;
}

body {
  background: var(--color-canvas);
  color: var(--color-text);
  font-family: var(--font-base);
  font-size: var(--font-size-md);
  line-height: var(--line-normal);
  -webkit-font-smoothing: antialiased;
}

h1, h2, h3, p, dl, dd, dt, ul, li { margin: 0; padding: 0; }
ul { list-style: none; }

:focus-visible {
  outline: var(--focus-ring) solid var(--color-focus);
  outline-offset: var(--focus-offset);
}

/* スクロール領域 ---------------------------------------------------- */
.scroll { overflow: auto; scrollbar-width: thin; scrollbar-color: var(--color-line) transparent; }
.scroll::-webkit-scrollbar { width: var(--scrollbar-size); height: var(--scrollbar-size); }
.scroll::-webkit-scrollbar-thumb {
  background: var(--color-line);
  border: var(--scrollbar-inset) solid transparent;
  background-clip: content-box;
  border-radius: var(--radius-pill);
}
.scroll::-webkit-scrollbar-thumb:hover { background: var(--color-surface-3); background-clip: content-box; }

/* 全体レイアウト（スクロールレス） ----------------------------------- */
.app {
  display: grid;
  grid-template-rows: var(--topbar-height) minmax(0, 1fr) var(--actionbar-height);
  height: 100dvh;
}

/* 上部バー（メインカラーの帯） -------------------------------------- */
.topbar {
  display: flex;
  align-items: center;
  gap: var(--space-5);
  padding: 0 var(--space-4);
  background: var(--color-main);
  border-bottom: var(--border-width) solid var(--color-main);
  color: var(--color-on-main);
}
.topbar :focus-visible { outline-color: var(--color-on-main); }
.topbar .btn {
  background: var(--color-on-main-quiet);
  border-color: var(--color-on-main-line);
  color: var(--color-on-main);
}
.topbar .btn:hover:not(:disabled) { background: var(--color-on-main-quiet-strong); }

.topbar__identity { display: flex; align-items: center; gap: var(--space-3); flex: 0 0 auto; }

.topbar__mark {
  display: grid;
  place-items: center;
  width: var(--mark-size);
  height: var(--mark-size);
  border-radius: var(--radius-md);
  background: var(--color-on-main-quiet-strong);
  border: var(--border-width) solid var(--color-on-main-line);
  color: var(--color-on-main);
  font-size: var(--font-size-xs);
  font-weight: 700;
  letter-spacing: 0.04em;
}

.topbar__title { font-size: var(--font-size-lg); font-weight: 600; line-height: var(--line-tight); color: var(--color-on-main); }
.topbar__sub { font-size: var(--font-size-xs); color: var(--color-on-main-muted); line-height: var(--line-tight); }

.facts {
  display: flex;
  align-items: center;
  gap: var(--space-2);
  flex: 1 1 auto;
  min-width: 0;
  overflow: hidden;
}

.fact {
  display: flex;
  align-items: baseline;
  gap: var(--space-2);
  min-width: 0;
  max-width: var(--fact-max-width);
  padding: var(--space-1) var(--space-3);
  border: var(--border-width) solid var(--color-on-main-line);
  border-radius: var(--radius-pill);
  background: var(--color-on-main-quiet);
}

.fact dt { flex: 0 0 auto; font-size: var(--font-size-xs); color: var(--color-on-main-muted); }
.fact dd {
  min-width: 0;
  font-family: var(--font-mono);
  font-size: var(--font-size-xs);
  color: var(--color-on-main);
  white-space: nowrap;
  overflow: hidden;
  text-overflow: ellipsis;
}
.fact--alert { border-color: var(--color-on-main-warn); background: var(--color-on-main-warn-quiet); }
.fact--alert dd { color: var(--color-on-main-warn); font-family: var(--font-base); }

.topbar__actions { display: flex; align-items: center; gap: var(--space-2); flex: 0 0 auto; }

/* ボタン ------------------------------------------------------------ */
.btn {
  display: inline-flex;
  align-items: center;
  justify-content: center;
  gap: var(--space-2);
  height: var(--control-height);
  padding: 0 var(--space-3);
  border: var(--border-width) solid var(--color-line);
  border-radius: var(--radius-sm);
  background: var(--color-surface-2);
  color: var(--color-text);
  flex: 0 0 auto;
  font-family: inherit;
  font-size: var(--font-size-sm);
  white-space: nowrap;
  cursor: pointer;
  transition: background var(--duration-fast) var(--easing), border-color var(--duration-fast) var(--easing);
}
.btn:hover:not(:disabled) { background: var(--color-surface-3); }
.btn:disabled { opacity: 0.45; cursor: not-allowed; }

.btn--quiet { background: transparent; }
.btn--quiet:hover:not(:disabled) { background: var(--color-surface-2); }

.btn--ghost { background: transparent; border-color: transparent; color: var(--color-text-muted); }
.btn--ghost:hover:not(:disabled) { background: var(--color-surface-2); color: var(--color-text); }

.btn--icon { width: var(--control-height); padding: 0; font-size: var(--font-size-lg); letter-spacing: 0.08em; }

.btn--primary {
  height: var(--control-height-lg);
  padding: 0 var(--space-5);
  background: var(--color-main);
  border-color: var(--color-main);
  color: var(--color-main-ink);
  font-size: var(--font-size-md);
  font-weight: 700;
}
.btn--primary:hover:not(:disabled) { background: var(--color-main-hover); border-color: var(--color-main-hover); }
.btn--primary:disabled { background: var(--color-surface-2); border-color: var(--color-line); color: var(--color-text-faint); }

.btn--danger { color: var(--color-danger); }
.btn--danger:hover:not(:disabled) { background: var(--color-danger-quiet); }

/* メニュー ---------------------------------------------------------- */
.menu { position: relative; }
.menu__list {
  position: absolute;
  top: calc(100% + var(--space-2));
  right: 0;
  z-index: 30;
  min-width: var(--menu-min-width);
  padding: var(--space-1);
  background: var(--color-surface);
  border: var(--border-width) solid var(--color-line);
  border-radius: var(--radius-md);
  box-shadow: var(--shadow-2);
}
.menu__list button {
  display: block;
  width: 100%;
  padding: var(--space-2) var(--space-3);
  border: 0;
  border-radius: var(--radius-sm);
  background: transparent;
  color: var(--color-text);
  font-family: inherit;
  font-size: var(--font-size-sm);
  text-align: left;
  cursor: pointer;
}
.menu__list button:hover { background: var(--color-surface-3); }
.menu__danger { color: var(--color-danger); }
.menu__split { height: 0; margin: var(--space-1) 0; border: 0; border-top: var(--border-width) solid var(--color-line-quiet); }

/* 終了後の画面 ------------------------------------------------------ */
.farewell {
  position: fixed;
  inset: 0;
  z-index: 60;
  display: grid;
  place-items: center;
  padding: var(--space-5);
  background: var(--color-canvas);
}
.farewell__box {
  max-width: 46ch;
  padding: var(--space-5);
  background: var(--color-surface);
  border: var(--border-width) solid var(--color-line);
  border-radius: var(--radius-lg);
  box-shadow: var(--shadow-2);
  text-align: center;
}
.farewell__title { font-size: var(--font-size-lg); font-weight: 600; color: var(--color-main); }
.farewell__text { margin-top: var(--space-2); font-size: var(--font-size-sm); color: var(--color-text-muted); }

/* 作業領域 ---------------------------------------------------------- */
.workspace {
  display: grid;
  grid-template-columns: minmax(var(--panel-min-rules), 40fr) minmax(var(--panel-min-result), 60fr);
  gap: var(--space-3);
  min-height: 0;
  padding: var(--space-3);
}

/* 見出し / （集計・取り消し） / 本体 / 脚注。本体だけが伸び縮みする */
.panel {
  display: flex;
  flex-direction: column;
  min-height: 0;
  background: var(--color-surface);
  border: var(--border-width) solid var(--color-line-quiet);
  border-radius: var(--radius-lg);
  overflow: hidden;
}
.panel > * { flex: 0 0 auto; }
.panel > .panel__body { flex: 1 1 auto; min-height: 0; }

.panel__head {
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: var(--space-3);
  padding: var(--space-3) var(--space-4);
  border-bottom: var(--border-width) solid var(--color-line-quiet);
}
.panel__title { font-size: var(--font-size-md); font-weight: 600; letter-spacing: 0.02em; color: var(--color-main); }
.panel__note { font-size: var(--font-size-xs); color: var(--color-text-faint); line-height: var(--line-tight); }
.panel__head-actions { display: flex; gap: var(--space-2); }
.panel__body { padding: var(--space-3); }
.panel__foot {
  padding: var(--space-2) var(--space-4);
  border-top: var(--border-width) solid var(--color-line-quiet);
  font-size: var(--font-size-xs);
  color: var(--color-text-faint);
}

/* ルールカード ------------------------------------------------------ */
.rule {
  margin-bottom: var(--space-3);
  background: var(--color-surface-2);
  border: var(--border-width) solid var(--color-line-quiet);
  border-radius: var(--radius-md);
  overflow: hidden;
}
.rule:last-child { margin-bottom: 0; }
.rule--disabled { opacity: 0.72; }
.rule--invalid { border-color: var(--color-danger); }

.rule__head {
  display: flex;
  align-items: center;
  gap: var(--space-2);
  padding: var(--space-2) var(--space-3);
  background: var(--color-well);
  border-bottom: var(--border-width) solid var(--color-line-quiet);
}
.rule__order {
  flex: 0 0 auto;
  display: flex;
  align-items: center;
  gap: var(--space-1);
}
.rule__order input {
  width: var(--input-width-priority);
  text-align: center;
}
.rule__moves { display: flex; flex-direction: column; gap: var(--space-hair); }
.rule__moves button {
  width: var(--stepper-width);
  height: var(--stepper-height);
  padding: 0;
  border: var(--border-width) solid var(--color-line);
  border-radius: var(--radius-xs);
  background: var(--color-surface-3);
  color: var(--color-text-muted);
  font-size: var(--font-size-2xs);
  line-height: 1;
  cursor: pointer;
}
.rule__moves button:disabled { opacity: 0.35; cursor: not-allowed; }
.rule__name { flex: 1 1 auto; min-width: 0; font-weight: 600; }

.rule__grid {
  display: grid;
  grid-template-columns: repeat(2, minmax(0, 1fr));
  gap: var(--space-2) var(--space-3);
  padding: var(--space-3);
}
.rule__grid .field--wide { grid-column: 1 / -1; }

.rule__foot {
  display: flex;
  align-items: baseline;
  gap: var(--space-2);
  padding: var(--space-2) var(--space-3);
  border-top: var(--border-width) solid var(--color-line-quiet);
  background: var(--color-well);
  font-size: var(--font-size-xs);
}
.rule__foot span { flex: 0 0 auto; color: var(--color-text-faint); }
.rule__foot code {
  min-width: 0;
  font-family: var(--font-mono);
  color: var(--color-text-muted);
  white-space: nowrap;
  overflow: hidden;
  text-overflow: ellipsis;
}

/* 入力部品 ---------------------------------------------------------- */
.field { display: flex; flex-direction: column; gap: var(--space-1); min-width: 0; }
.field label, .field > .field__label {
  font-size: var(--font-size-xs);
  color: var(--color-text-muted);
}
.field__hint { font-size: var(--font-size-xs); color: var(--color-text-faint); line-height: var(--line-tight); }
.field__error {
  display: flex;
  gap: var(--space-1);
  font-size: var(--font-size-xs);
  color: var(--color-danger);
  line-height: var(--line-tight);
}
.field__unit { display: flex; align-items: center; gap: var(--space-2); }
.field__unit input[type="number"] { width: var(--input-width-number); flex: 0 0 auto; }
.field__unit span { flex: 0 0 auto; font-size: var(--font-size-sm); color: var(--color-text-muted); white-space: nowrap; }

input[type="text"], input[type="number"], select {
  width: 100%;
  height: var(--control-height);
  padding: 0 var(--space-2);
  border: var(--border-width) solid var(--color-line);
  border-radius: var(--radius-sm);
  background: var(--color-field);
  color: var(--color-text);
  font-family: inherit;
  font-size: var(--font-size-sm);
}
input::placeholder { color: var(--color-placeholder); }
input:hover, select:hover { border-color: var(--color-line); }
input:focus, select:focus { border-color: var(--color-accent); }
input[aria-invalid="true"] { border-color: var(--color-danger); }
select { cursor: pointer; }
option { background: var(--color-field); color: var(--color-text); }

.check { display: inline-flex; align-items: center; gap: var(--space-2); flex: 0 0 auto; cursor: pointer; }
.check input { position: absolute; opacity: 0; width: 0; height: 0; }
.check__box {
  position: relative;
  width: var(--switch-width);
  height: var(--switch-height);
  border-radius: var(--radius-pill);
  background: var(--color-surface-3);
  border: var(--border-width) solid var(--color-line);
  transition: background var(--duration-fast) var(--easing);
}
.check__box::after {
  content: "";
  position: absolute;
  top: var(--space-0);
  left: var(--space-0);
  width: var(--switch-knob);
  height: var(--switch-knob);
  border-radius: 50%;
  background: var(--color-text-faint);
  transition: transform var(--duration-fast) var(--easing), background var(--duration-fast) var(--easing);
}
.check input:checked + .check__box { background: var(--color-accent-quiet); border-color: var(--color-accent); }
.check input:checked + .check__box::after { transform: translateX(var(--switch-travel)); background: var(--color-accent); }
.check input:focus-visible + .check__box { outline: var(--focus-ring) solid var(--color-focus); outline-offset: var(--focus-offset); }
.check__text { font-size: var(--font-size-sm); color: var(--color-text); white-space: nowrap; }
.rule__head .check__text { min-width: 2.4em; }

/* 状態チップ（アイコン＋文字＋色の三重表現） -------------------------- */
.chip {
  display: inline-flex;
  align-items: center;
  gap: var(--space-1);
  padding: var(--space-hair) var(--space-2);
  border: var(--border-width) solid transparent;
  border-radius: var(--radius-pill);
  font-size: var(--font-size-xs);
  white-space: nowrap;
}
.chip__icon { font-size: var(--font-size-icon); }
.chip--success { color: var(--color-success); background: var(--color-success-quiet); border-color: var(--color-success-edge); }
.chip--warn { color: var(--color-warn); background: var(--color-warn-quiet); border-color: var(--color-warn-edge); }
.chip--danger { color: var(--color-danger); background: var(--color-danger-quiet); border-color: var(--color-danger-edge); }
.chip--info { color: var(--color-info); background: var(--color-info-quiet); border-color: var(--color-info-edge); }
.chip--neutral { color: var(--color-neutral); background: var(--color-neutral-quiet); border-color: var(--color-neutral-edge); }

/* 集計と絞り込み ---------------------------------------------------- */
.summary {
  display: flex;
  flex-wrap: wrap;
  gap: var(--space-2);
  padding: var(--space-2) var(--space-4);
  border-bottom: var(--border-width) solid var(--color-line-quiet);
}
.summary__item {
  display: inline-flex;
  align-items: center;
  gap: var(--space-2);
  padding: var(--space-1) var(--space-3);
  border: var(--border-width) solid var(--color-line-quiet);
  border-radius: var(--radius-pill);
  background: transparent;
  color: var(--color-text-muted);
  font-family: inherit;
  font-size: var(--font-size-xs);
  cursor: pointer;
}
.summary__item:hover { background: var(--color-surface-2); }
.summary__item[aria-pressed="true"] { border-color: var(--color-accent); background: var(--color-accent-quiet); color: var(--color-text); }
.summary__item strong { font-size: var(--font-size-md); font-variant-numeric: tabular-nums; }
.summary__item--zero { opacity: 0.5; }
.summary__filter { margin-left: auto; display: flex; align-items: center; gap: var(--space-2); font-size: var(--font-size-xs); color: var(--color-text-muted); }

/* 取り消しバー ------------------------------------------------------ */
.undo-bar {
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: var(--space-3);
  padding: var(--space-2) var(--space-4);
  background: var(--color-accent-quiet);
  border-bottom: var(--border-width) solid var(--color-accent-edge);
  font-size: var(--font-size-sm);
}

/* 結果テーブル ------------------------------------------------------ */
.table { width: 100%; table-layout: fixed; border-collapse: collapse; font-size: var(--font-size-sm); }
.table col.col-result { width: 12%; }
.table col.col-name { width: 24%; }
.table col.col-month { width: 8%; }
.table col.col-rule { width: 14%; }
.table col.col-planned { width: 13%; }
.table col.col-actual { width: 13%; }
.table col.col-message { width: 16%; }
.table th {
  position: sticky;
  top: calc(var(--space-3) * -1);
  z-index: 1;
  padding: var(--space-2);
  background: var(--color-surface-2);
  border-bottom: var(--border-width) solid var(--color-line);
  color: var(--color-text-muted);
  font-size: var(--font-size-xs);
  font-weight: 600;
  text-align: left;
  white-space: nowrap;
}
.table td {
  padding: var(--space-2);
  border-bottom: var(--border-width) solid var(--color-line-quiet);
  vertical-align: top;
}
.table tbody tr:hover { background: var(--color-row-hover); }
.table .cell-path, .table .cell-name { font-family: var(--font-mono); font-size: var(--font-size-xs); }
.table .cell-path { color: var(--color-text-muted); }
.table .truncate {
  display: block;
  max-width: 100%;
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
}
.table .cell-name .truncate { color: var(--color-text); }
.table .cell-month { font-variant-numeric: tabular-nums; white-space: nowrap; }
.table .cell-message { color: var(--color-text-muted); font-size: var(--font-size-xs); }
.row-toggle {
  width: 100%;
  padding: 0;
  border: 0;
  background: transparent;
  color: inherit;
  font: inherit;
  text-align: left;
  cursor: pointer;
}
.detail td { background: var(--color-well-strong); }
.detail dl { display: grid; grid-template-columns: 7em 1fr; gap: var(--space-1) var(--space-3); }
.detail dt { font-size: var(--font-size-xs); color: var(--color-text-faint); }
.detail dd { font-family: var(--font-mono); font-size: var(--font-size-xs); color: var(--color-text-muted); word-break: break-all; }
.detail .reasons { display: flex; flex-direction: column; gap: var(--space-0); font-family: var(--font-base); }

/* 対象ファイル一覧（判定前） ----------------------------------------- */
.filelist { display: flex; flex-direction: column; gap: var(--space-1); }
.filelist__row {
  display: grid;
  grid-template-columns: minmax(0, 1fr) auto auto;
  align-items: center;
  gap: var(--space-3);
  padding: var(--space-2) var(--space-3);
  background: var(--color-surface-2);
  border: var(--border-width) solid var(--color-line-quiet);
  border-radius: var(--radius-sm);
}
.filelist__name { font-family: var(--font-mono); font-size: var(--font-size-xs); overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.filelist__meta { font-size: var(--font-size-xs); color: var(--color-text-faint); white-space: nowrap; }

/* 空の状態 ---------------------------------------------------------- */
.empty {
  display: grid;
  place-content: center;
  justify-items: center;
  gap: var(--space-2);
  height: 100%;
  padding: var(--space-5);
  text-align: center;
}
.empty__title { font-size: var(--font-size-md); color: var(--color-text-muted); }
.empty__text { font-size: var(--font-size-sm); color: var(--color-text-faint); max-width: 46ch; }
.empty__path {
  max-width: 100%;
  padding: var(--space-1) var(--space-3);
  border: var(--border-width) solid var(--color-line-quiet);
  border-radius: var(--radius-sm);
  background: var(--color-well-strong);
  font-family: var(--font-mono);
  font-size: var(--font-size-xs);
  color: var(--color-text-muted);
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
}

/* 下部アクションバー ------------------------------------------------- */
.actionbar {
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: var(--space-4);
  padding: 0 var(--space-4);
  background: var(--color-bar-shade);
  border-top: var(--border-width) solid var(--color-line);
}
.actionbar__status { min-width: 0; }
.actionbar__reason,
.actionbar__hint {
  overflow: hidden;
  white-space: nowrap;
  text-overflow: ellipsis;
  line-height: var(--line-tight);
}
.actionbar__reason { font-size: var(--font-size-sm); }
.actionbar__hint { font-size: var(--font-size-xs); color: var(--color-text-faint); }
.actionbar__reason .chip { margin-right: var(--space-2); }
.actionbar__buttons { display: flex; align-items: center; gap: var(--space-2); flex: 0 0 auto; }

/* トースト ---------------------------------------------------------- */
.toasts {
  position: fixed;
  left: var(--space-4);
  bottom: calc(var(--actionbar-height) + var(--space-3));
  z-index: 50;
  display: flex;
  flex-direction: column;
  gap: var(--space-2);
}
.toast {
  display: flex;
  align-items: center;
  gap: var(--space-3);
  padding: var(--space-2) var(--space-3);
  background: var(--color-surface);
  border: var(--border-width) solid var(--color-line);
  border-radius: var(--radius-md);
  box-shadow: var(--shadow-2);
  font-size: var(--font-size-sm);
  animation: toast-in var(--duration) var(--easing);
}
@keyframes toast-in { from { opacity: 0; transform: translateY(var(--lift)); } to { opacity: 1; transform: none; } }

/* ダイアログ -------------------------------------------------------- */
.dialog {
  width: min(var(--dialog-width), 92vw);
  max-height: 86dvh;
  padding: 0;
  border: var(--border-width) solid var(--color-line);
  border-radius: var(--radius-lg);
  background: var(--color-surface);
  color: var(--color-text);
  box-shadow: var(--shadow-2);
}
.dialog::backdrop { background: var(--color-backdrop); }
.dialog__form { display: flex; flex-direction: column; max-height: 86dvh; }
.dialog__head { padding: var(--space-4); border-bottom: var(--border-width) solid var(--color-line-quiet); }
.dialog__head h2 { font-size: var(--font-size-lg); font-weight: 600; color: var(--color-main); }
.dialog__head p { margin-top: var(--space-1); font-size: var(--font-size-xs); color: var(--color-text-faint); }
.dialog__lead {
  margin-top: var(--space-3);
  padding: var(--space-2) var(--space-3);
  border: var(--border-width) solid var(--color-accent-edge);
  border-radius: var(--radius-sm);
  background: var(--color-accent-quiet);
  font-size: var(--font-size-sm);
  color: var(--color-text);
}
.dialog__body {
  display: grid;
  grid-template-columns: repeat(2, minmax(0, 1fr));
  gap: var(--space-4);
  padding: var(--space-4);
}
.dialog__body .field--wide { grid-column: 1 / -1; }
.dialog__foot {
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: var(--space-3);
  padding: var(--space-3) var(--space-4);
  border-top: var(--border-width) solid var(--color-line-quiet);
}
.dialog__foot-note { font-size: var(--font-size-xs); color: var(--color-text-faint); }

.fixed-list {
  padding: var(--space-3);
  border: var(--border-width) dashed var(--color-line);
  border-radius: var(--radius-md);
  background: var(--color-well);
}
.fixed-list h3 { font-size: var(--font-size-xs); color: var(--color-text-muted); margin-bottom: var(--space-2); }
.fixed-list li {
  display: flex;
  gap: var(--space-2);
  font-size: var(--font-size-xs);
  color: var(--color-text-faint);
  line-height: var(--line-tight);
  padding: var(--space-0) 0;
}
.fixed-list b { color: var(--color-text-muted); font-weight: 600; flex: 0 0 12em; }

/* 画面が狭い場合 ---------------------------------------------------- */
@media (max-width: 1100px) {
  .facts { display: none; }
}
@media (max-width: 900px) {
  .workspace { grid-template-columns: minmax(0, 1fr); grid-template-rows: minmax(0, 1fr) minmax(0, 1fr); }
  .rule__grid { grid-template-columns: minmax(0, 1fr); }
  .dialog__body { grid-template-columns: minmax(0, 1fr); }
}

@media (prefers-reduced-motion: reduce) {
  * { animation-duration: 1ms !important; transition-duration: 1ms !important; }
}
"""


UI_JS = r"""'use strict';

/* =======================================================================
   状態
   ===================================================================== */

const SETTING_KEYS = [
  'common_base', 'unknown_folder', 'log_folder', 'use_month_folder', 'month_folder_format',
  'month_source', 'month_fallback', 'duplicate_mode', 'evacuation_folder_prefix', 'forbid_overwrite',
  'move_strategy', 'verify_after_move', 'check_file_stable', 'min_file_age_seconds',
  'backup_rules_on_save', 'max_path_length',
];

const state = {
  settings: {},
  rules: [],
  meta: null,
  baseCheckedFor: null,
  fatal: false,
  dirty: false,
  busy: false,
  view: { mode: 'idle', items: [], summary: null, at: '', filter: null },
  expanded: new Set(),
  nextId: 1,
  closed: false,
};

let firstLoad = true;

const RESULT_INFO = {
  MATCH: { label: '移動予定', icon: '→', tone: 'info' },
  MOVED: { label: '移動成功', icon: '✔', tone: 'success' },
  COPIED_AND_DELETED: { label: '移動成功', icon: '✔', tone: 'success' },
  COPIED: { label: 'コピー', icon: '✔', tone: 'success' },
  DUPLICATE_EVACUATED: { label: '避難保存', icon: '▲', tone: 'warn' },
  DUPLICATE_WILL_EVACUATE: { label: '避難保存', icon: '▲', tone: 'warn' },
  UNKNOWN: { label: '判定不能', icon: '?', tone: 'neutral' },
  MONTH_UNKNOWN: { label: '年月不明', icon: '▲', tone: 'warn' },
  MONTH_AMBIGUOUS: { label: '年月確認要', icon: '▲', tone: 'warn' },
  SKIP: { label: 'スキップ', icon: '—', tone: 'neutral' },
  SKIP_FILE_NOT_STABLE: { label: 'スキップ', icon: '—', tone: 'neutral' },
  UNDONE: { label: '取り消し済み', icon: '↩', tone: 'success' },
};

const PREVIEW_LABEL = {
  MOVED: '移動予定', COPIED_AND_DELETED: '移動予定', MATCH: '移動予定',
  DUPLICATE_EVACUATED: '避難予定',
};

const CATEGORIES = [
  { key: 'success', label: '移動成功', previewLabel: '移動予定', icon: '✔', tone: 'success' },
  { key: 'evacuated', label: '避難保存', previewLabel: '避難予定', icon: '▲', tone: 'warn' },
  { key: 'month', label: '年月退避', previewLabel: '年月退避予定', icon: '▲', tone: 'warn' },
  { key: 'unknown', label: '判定不能', previewLabel: '判定不能', icon: '?', tone: 'neutral' },
  { key: 'error', label: 'エラー', previewLabel: 'エラー', icon: '✕', tone: 'danger' },
  { key: 'skip', label: 'スキップ', previewLabel: 'スキップ', icon: '—', tone: 'neutral' },
];

const FORBIDDEN_CHARS = ['<', '>', ':', '"', '|', '?', '*'];
const RESERVED_NAMES = new Set(['CON', 'PRN', 'AUX', 'NUL',
  'COM1', 'COM2', 'COM3', 'COM4', 'COM5', 'COM6', 'COM7', 'COM8', 'COM9',
  'LPT1', 'LPT2', 'LPT3', 'LPT4', 'LPT5', 'LPT6', 'LPT7', 'LPT8', 'LPT9']);

/* =======================================================================
   小さな道具
   ===================================================================== */

const $ = (selector) => document.querySelector(selector);

function esc(value) {
  return String(value === null || value === undefined ? '' : value)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;').replace(/'/g, '&#39;');
}

function splitKeywords(text) {
  return String(text || '').split(/[,、\n\r\t]/).map((item) => item.trim()).filter(Boolean);
}

function joinKeywords(list) {
  return (list || []).join(', ');
}

function categorize(result) {
  if (result === 'MOVED' || result === 'COPIED_AND_DELETED' || result === 'COPIED'
    || result === 'MATCH' || result === 'UNDONE') return 'success';
  if (result === 'DUPLICATE_EVACUATED' || result === 'DUPLICATE_WILL_EVACUATE') return 'evacuated';
  if (result === 'MONTH_UNKNOWN' || result === 'MONTH_AMBIGUOUS') return 'month';
  if (result === 'UNKNOWN') return 'unknown';
  if (result.startsWith('ERROR_')) return 'error';
  return 'skip';
}

function resultChip(result, mode) {
  const info = RESULT_INFO[result] || { label: 'エラー', icon: '✕', tone: 'danger' };
  const label = mode === 'preview' && PREVIEW_LABEL[result] ? PREVIEW_LABEL[result] : info.label;
  return `<span class="chip chip--${info.tone}"><span class="chip__icon" aria-hidden="true">${info.icon}</span>${esc(label)}</span>`;
}

function shortenPath(path) {
  if (!path) return '';
  const base = String(state.settings.common_base || '');
  if (base && base !== '{BASE_FOLDER}' && path.startsWith(base) && path.length > base.length) {
    return '…' + path.slice(base.length);
  }
  const script = state.meta ? state.meta.script_folder : '';
  if (script && path.startsWith(script) && path.length > script.length) return '…' + path.slice(script.length);
  return path;
}

/* 一覧では保存先「フォルダ」を見せる（ファイル名は専用の列にある） */
function folderOf(path) {
  if (!path) return '';
  const index = Math.max(path.lastIndexOf('\\'), path.lastIndexOf('/'));
  return index > 0 ? path.slice(0, index) : path;
}

/* 長いパスは末尾（フォルダ名側）を残して省略する */
function tailText(text, limit) {
  const value = String(text || '');
  return value.length > limit ? '…' + value.slice(value.length - limit) : value;
}

function pathSeparator(sample) {
  const text = String(sample || state.settings.common_base || '');
  return text.includes('\\') || /^[A-Za-z]:/.test(text) ? '\\' : '/';
}

function joinPath(...parts) {
  const sep = pathSeparator(parts[0]);
  return parts.filter(Boolean).join(sep).replace(/[\\/]+$/, '');
}

/* ルールが使う保存先フォルダ（未指定なら共通保存先） */
function ruleBase(rule) {
  const own = String(rule.destination_base || '').trim();
  return { path: own || String(state.settings.common_base || '').trim(), own: !!own };
}

/* 共通保存先を使う有効なルールがあるか */
function commonBaseRequired() {
  return state.rules.some((rule) => rule.enabled && !String(rule.destination_base || '').trim());
}

function formatBytes(size) {
  if (size < 1024) return `${size} バイト`;
  if (size < 1024 * 1024) return `${(size / 1024).toFixed(1)} KB`;
  return `${(size / 1024 / 1024).toFixed(1)} MB`;
}

/* =======================================================================
   通信
   ===================================================================== */

async function api(path, method, body) {
  if (state.closed) return { ok: false, status: 0, data: { ok: false, message: 'ツールは終了しています' } };
  const options = { method: method || 'GET', headers: { 'Accept': 'application/json' } };
  if (body !== undefined && body !== null) {
    options.headers['Content-Type'] = 'application/json';
    options.body = JSON.stringify(body);
  }
  let response;
  try {
    response = await fetch(path, options);
  } catch (error) {
    return { ok: false, status: 0, data: { ok: false, message: `画面とツールの通信ができません（${error.message}）。コンソールウィンドウが閉じていないか確認してください。` } };
  }
  let data;
  try {
    data = await response.json();
  } catch {
    data = { ok: false, message: '応答を解釈できませんでした' };
  }
  return { ok: response.ok, status: response.status, data };
}

function currentPayload() {
  const payload = {};
  SETTING_KEYS.forEach((key) => { payload[key] = state.settings[key]; });
  payload.rules = state.rules.map((rule) => ({
    enabled: !!rule.enabled,
    priority: Number(rule.priority) || 0,
    name: rule.name || '',
    extension: rule.extension || '.pdf',
    contains_all: splitKeywords(rule.contains_all_text),
    contains_any: splitKeywords(rule.contains_any_text),
    not_contains: splitKeywords(rule.not_contains_text),
    destination_base: (rule.destination_base || '').trim(),
    destination_subfolder: rule.destination_subfolder || '',
  }));
  return payload;
}

function adoptSettings(payload) {
  state.settings = {};
  SETTING_KEYS.forEach((key) => { state.settings[key] = payload[key]; });
  state.rules = (payload.rules || []).map((rule) => ({
    _id: state.nextId++,
    enabled: !!rule.enabled,
    priority: rule.priority,
    name: rule.name || '',
    extension: rule.extension || '.pdf',
    contains_all_text: joinKeywords(rule.contains_all),
    contains_any_text: joinKeywords(rule.contains_any),
    not_contains_text: joinKeywords(rule.not_contains),
    destination_base: rule.destination_base || '',
    destination_subfolder: rule.destination_subfolder || '',
  }));
  sortRules();
  state.dirty = false;
}

function adoptMeta(meta) {
  if (!meta) return;
  state.meta = meta;
  state.baseCheckedFor = state.settings.common_base;
}

async function loadAll(quiet) {
  const { data } = await api('/api/rules');
  if (!data || data.ok === false) {
    toast(data && data.message ? data.message : '設定を読み込めませんでした');
    return;
  }
  adoptSettings(data);
  state.fatal = !!data._fatal;
  adoptMeta(data._meta);
  state.view = { mode: 'idle', items: [], summary: null, at: '', filter: null };
  renderAll();
  if (data._errors && data._errors.length) {
    toast(`rules.json に ${data._errors.length} 件の問題があります。内容を確認してください。`);
  } else if (!quiet) {
    toast('rules.json を読み込みました');
  }

  if (firstLoad) {
    firstLoad = false;
    const base = String(state.settings.common_base || '').trim();
    if ((!base || base === '{BASE_FOLDER}') && !state.rules.length) {
      bindSettingsDialog();
      $('#settings-dialog').showModal();
    }
  }
}

async function refreshMeta() {
  const { data } = await api('/api/status');
  if (data && data._meta) {
    const keepBase = state.baseCheckedFor;
    state.meta = data._meta;
    state.baseCheckedFor = keepBase;
    renderFacts();
    renderActionBar();
    renderUndoBar();
    if (state.view.mode === 'idle') renderResults();
  }
}

/* =======================================================================
   検証（画面側。サーバー側の検証と同じ規則）
   ===================================================================== */

function segmentError(segment, label) {
  if (segment === '.' || segment === '..') return `${label}に ${segment} は使用できません`;
  if (segment !== segment.replace(/[ .]+$/, '')) return `${label}の末尾に空白またはピリオドは使用できません`;
  if (RESERVED_NAMES.has(segment.split('.')[0].toUpperCase())) return `${label}に Windows の予約語 ${segment} は使用できません`;
  return '';
}

function folderNameError(name, label) {
  if (!name) return `${label}を入力してください`;
  for (const char of FORBIDDEN_CHARS.concat(['/', '\\'])) {
    if (name.includes(char)) return `${label}に使用できない文字 ${char} が含まれています`;
  }
  return segmentError(name, label);
}

function baseFormatError(value, label) {
  if (!value) return '';
  for (const char of FORBIDDEN_CHARS) {
    if (char !== ':' && value.includes(char)) return `${label}に使用できない文字 ${char} が含まれています`;
  }
  if (value.split(/[\\/]/).includes('..')) return `${label}に .. は使用できません`;
  if (!/^([A-Za-z]:[\\/]|\\\\|\/)/.test(value)) {
    return `${label}は絶対パスで指定してください（例: D:\\共有\\設備記録）`;
  }
  return '';
}

function subfolderError(value, label) {
  if (!value) return `${label}を入力してください`;
  if (/^[A-Za-z]:/.test(value) || value.startsWith('/') || value.startsWith('\\')) {
    return `${label}に絶対パスは指定できません（共通保存先からの相対パスで指定してください）`;
  }
  for (const char of FORBIDDEN_CHARS) {
    if (value.includes(char)) return `${label}に使用できない文字 ${char} が含まれています`;
  }
  const segments = value.split(/[\\/]/);
  if (segments.some((segment) => segment === '')) return `${label}に空のフォルダ名が含まれています`;
  for (const segment of segments) {
    const error = segmentError(segment, label);
    if (error) return error;
  }
  return '';
}

function validate() {
  const ruleErrors = new Map();
  const messages = [];

  state.rules.forEach((rule, index) => {
    const errors = {};
    const label = `ルール${index + 1}`;
    if (!String(rule.name || '').trim()) errors.name = 'ルール名を入力してください';
    const extension = String(rule.extension || '').trim().toLowerCase();
    if (extension && extension !== '.pdf') errors.extension = '初期仕様の対象は .pdf のみです';
    if (!splitKeywords(rule.contains_all_text).length && !splitKeywords(rule.contains_any_text).length) {
      errors.contains_all_text = 'すべて含む／いずれか含む のどちらかにキーワードが必要です';
    }
    const ownBaseError = baseFormatError(String(rule.destination_base || '').trim(), '保存先フォルダ');
    if (ownBaseError) errors.destination_base = ownBaseError;
    const subError = subfolderError(String(rule.destination_subfolder || '').trim(), '保存先サブフォルダ');
    if (subError) errors.destination_subfolder = subError;

    if (Object.keys(errors).length) {
      ruleErrors.set(rule._id, errors);
      Object.values(errors).forEach((message) => messages.push(`${label}: ${message}`));
    }
  });

  const unknownError = folderNameError(String(state.settings.unknown_folder || '').trim(), '判定不能フォルダ名');
  if (unknownError) messages.push(unknownError);
  const logError = folderNameError(String(state.settings.log_folder || '').trim(), 'ログフォルダ名');
  if (logError) messages.push(logError);

  const seconds = Number(state.settings.min_file_age_seconds);
  if (!Number.isFinite(seconds) || seconds < 0 || seconds > 3600) messages.push('最小ファイル経過秒数は 0〜3600 秒で指定してください');
  const pathLimit = Number(state.settings.max_path_length);
  if (!Number.isFinite(pathLimit) || pathLimit < 60 || pathLimit > 32767) messages.push('最大パス長は 60〜32767 文字で指定してください');

  return { ruleErrors, messages };
}

function baseError() {
  if (!commonBaseRequired()) return '';
  const base = String(state.settings.common_base || '').trim();
  if (!base || base === '{BASE_FOLDER}') return '共通保存先フォルダが未設定です';
  for (const char of ['<', '>', '"', '|', '?', '*']) {
    if (base.includes(char)) return `共通保存先フォルダに使用できない文字 ${char} が含まれています`;
  }
  if (!/^([A-Za-z]:[\\/]|\\\\|\/)/.test(base)) return '共通保存先フォルダは絶対パスで指定してください';
  if (state.meta && state.baseCheckedFor === state.settings.common_base && state.meta.base_folder_error) {
    return state.meta.base_folder_error;
  }
  return '';
}

/* =======================================================================
   描画
   ===================================================================== */

function renderAll() {
  renderFacts();
  renderRules();
  renderResults();
  renderSummary();
  renderUndoBar();
  renderActionBar();
  bindSettingsDialog();
}

function renderFacts() {
  const meta = state.meta;
  const facts = $('#facts');
  if (!meta) { facts.innerHTML = ''; return; }

  const base = String(state.settings.common_base || '').trim();
  const baseProblem = baseError();
  const hasBase = base && base !== '{BASE_FOLDER}';
  const parts = [];
  parts.push(fact('作業フォルダ', meta.script_folder, false));
  parts.push(fact('共通保存先',
    baseProblem ? baseProblem : (hasBase ? base : 'ルールごとに指定'), !!baseProblem));
  parts.push(fact('ログ', meta.log_path, false));
  facts.innerHTML = parts.join('');

  $('#topbar-sub').textContent =
    `このフォルダ直下の ${meta.target_extension} ${meta.file_count} 件が対象です（${meta.generated_at} 時点）`;
}

function fact(term, description, alert) {
  return `<div class="fact${alert ? ' fact--alert' : ''}">
    <dt>${esc(term)}</dt>
    <dd title="${esc(description)}">${esc(alert ? description : tailText(description, 26))}</dd>
  </div>`;
}

function sortRules() {
  state.rules.sort((a, b) => (Number(a.priority) || 0) - (Number(b.priority) || 0));
}

function renderRules() {
  const list = $('#rule-list');
  const validation = validate();
  const enabledCount = state.rules.filter((rule) => rule.enabled).length;

  $('#rules-note').textContent = state.rules.length
    ? `${state.rules.length} 件中 ${enabledCount} 件が有効`
    : 'ルールがありません';

  if (!state.rules.length) {
    list.innerHTML = `<div class="empty">
      <p class="empty__title">ルールがまだありません</p>
      <p class="empty__text">ルールを追加すると、ファイル名のキーワードでPDFの保存先を決められます。
      ルールが 0 件のときは、すべてのPDFが判定不能フォルダへ移動します。</p>
    </div>`;
    return;
  }

  list.innerHTML = state.rules.map((rule, index) => ruleCard(rule, index, validation.ruleErrors.get(rule._id) || {})).join('');
}

function ruleCard(rule, index, errors) {
  const invalid = Object.keys(errors).length > 0;
  const classes = ['rule'];
  if (!rule.enabled) classes.push('rule--disabled');
  if (invalid) classes.push('rule--invalid');

  return `<article class="${classes.join(' ')}" data-id="${rule._id}">
    <header class="rule__head">
      <label class="check" title="このルールを判定に使うかどうか">
        <input type="checkbox" data-field="enabled" ${rule.enabled ? 'checked' : ''}>
        <span class="check__box" aria-hidden="true"></span>
        <span class="check__text">${rule.enabled ? '有効' : '無効'}</span>
      </label>
      <div class="rule__order">
        <div class="rule__moves">
          <button type="button" data-action="up" aria-label="優先順位を上げる" ${index === 0 ? 'disabled' : ''}>▲</button>
          <button type="button" data-action="down" aria-label="優先順位を下げる" ${index === state.rules.length - 1 ? 'disabled' : ''}>▼</button>
        </div>
        <input type="number" data-field="priority" value="${esc(rule.priority)}" min="0" max="99999"
               aria-label="優先順位" title="小さい値ほど先に判定します">
      </div>
      <input type="text" class="rule__name" data-field="name" value="${esc(rule.name)}"
             aria-label="ルール名" placeholder="ルール名（例: 設備A 校正記録）"
             ${errors.name ? 'aria-invalid="true"' : ''}>
      <button type="button" class="btn btn--ghost" data-action="duplicate">複製</button>
      <button type="button" class="btn btn--ghost btn--danger" data-action="delete">削除</button>
    </header>

    <div class="rule__grid">
      ${field('すべて含むキーワード', 'contains_all_text', rule.contains_all_text, errors.contains_all_text,
    'カンマ区切り。すべて含む場合に一致します', true)}
      ${field('いずれか含むキーワード', 'contains_any_text', rule.contains_any_text, null,
    '任意。1つでも含めば条件を満たします', false)}
      ${field('含んではいけないキーワード', 'not_contains_text', rule.not_contains_text, null,
    '任意。1つでも含む場合は一致しません', false)}
      ${field('このルール専用の保存先フォルダ', 'destination_base', rule.destination_base, errors.destination_base,
    '空欄なら共通保存先を使います（例: D:\\共有\\設備A）', true)}
      ${field('保存先サブフォルダ', 'destination_subfolder', rule.destination_subfolder, errors.destination_subfolder,
    '保存先フォルダの下に作るフォルダ名', false)}
      ${field('拡張子', 'extension', rule.extension, errors.extension, '初期仕様では .pdf のみ', false)}
    </div>

    <footer class="rule__foot">
      <span>保存先</span>
      <code title="${esc(destinationExample(rule, true))}">${esc(destinationExample(rule, false))}</code>
      ${String(rule.destination_base || '').trim()
    ? '<span class="chip chip--info"><span class="chip__icon" aria-hidden="true">→</span>このルール専用</span>'
    : ''}
    </footer>
  </article>`;
}

function field(label, name, value, error, hint, wide) {
  const id = `f-${name}-${Math.random().toString(36).slice(2, 8)}`;
  return `<div class="field${wide ? ' field--wide' : ''}">
    <label for="${id}">${esc(label)}</label>
    <input type="text" id="${id}" data-field="${name}" value="${esc(value)}" spellcheck="false"
           ${error ? 'aria-invalid="true"' : ''}>
    ${error
      ? `<p class="field__error"><span aria-hidden="true">✕</span>${esc(error)}</p>`
      : `<p class="field__hint">${esc(hint)}</p>`}
  </div>`;
}

function destinationExample(rule, full) {
  const { path: base, own } = ruleBase(rule);
  if (!base || base === '{BASE_FOLDER}') {
    return own ? '保存先フォルダが未入力です' : '共通保存先フォルダが未設定です（このルール専用の保存先でも指定できます）';
  }
  const sub = String(rule.destination_subfolder || '').trim();
  if (!sub) return '保存先サブフォルダが未入力です';
  const month = state.settings.use_month_folder ? 'YYYY-MM' : '';
  const path = joinPath(base, sub, month);
  if (full) return path;
  return own ? tailText(path, 44) : shortenPath(path);
}

function renderSummary() {
  const box = $('#summary');
  const view = state.view;
  // 取り消し結果は「移動成功／避難保存…」の分類が当てはまらないため集計を出さない
  if (view.mode === 'idle' || view.mode === 'undone' || !view.summary) { box.hidden = true; return; }

  box.hidden = false;
  const chips = CATEGORIES.map((category) => {
    const count = view.summary[category.key] || 0;
    const active = view.filter === category.key;
    return `<button type="button" class="summary__item${count === 0 ? ' summary__item--zero' : ''}"
      data-filter="${category.key}" aria-pressed="${active}"
      title="${count === 0 ? '該当がないため絞り込めません' : 'クリックでこの結果だけ表示'}" ${count === 0 ? 'disabled' : ''}>
      <span class="chip chip--${category.tone}"><span class="chip__icon" aria-hidden="true">${category.icon}</span>${esc(view.mode === 'preview' ? category.previewLabel : category.label)}</span>
      <strong>${count}</strong> 件
    </button>`;
  }).join('');

  const filterNote = view.filter
    ? `<div class="summary__filter"><span>絞り込み中</span>
        <button type="button" class="btn btn--ghost" data-filter="clear">すべて表示</button></div>`
    : '';
  box.innerHTML = chips + filterNote;
}

function renderUndoBar() {
  const bar = $('#undo-bar');
  const undo = state.meta && state.meta.undo;
  if (!undo || !undo.available) { bar.hidden = true; return; }
  bar.hidden = false;
  bar.innerHTML = `<span>${undo.count} 件の移動を元に戻せます（${esc(undo.started_at)} の実行分）</span>
    <button type="button" class="btn btn--quiet" id="btn-undo" ${state.busy ? 'disabled' : ''}>移動を取り消す</button>`;
}

function renderResults() {
  const area = $('#result-area');
  const note = $('#result-note');
  const headActions = $('#result-head-actions');
  const view = state.view;

  if (view.mode === 'idle') {
    headActions.innerHTML = '';
    const files = (state.meta && state.meta.files) || [];
    note.textContent = files.length
      ? `このフォルダにある対象PDF ${files.length} 件（まだ判定していません）`
      : '対象PDFがありません';
    const folder = state.meta ? state.meta.script_folder : 'このツールと同じフォルダ';
    area.innerHTML = files.length ? fileList(files) : `<div class="empty">
        <p class="empty__title">対象のPDFがありません</p>
        <p class="empty__text">仕分けしたいPDFを、このツールと同じフォルダに置いてから「再読み込み」を押してください。
        サブフォルダ内のPDFと一時ファイルは対象外です。</p>
        <p class="empty__path" title="${esc(folder)}">${esc(tailText(folder, 52))}</p>
      </div>`;
    return;
  }

  const modeLabel = view.mode === 'preview' ? 'プレビュー結果（ファイルは移動していません）'
    : view.mode === 'undone' ? '取り消し結果' : '実行結果';
  note.textContent = `${view.at} 時点の${modeLabel}`;
  headActions.innerHTML = `<button type="button" class="btn btn--ghost" data-action="back-to-files">対象PDF一覧へ戻る</button>`;

  const items = view.filter ? view.items.filter((item) => categorize(item.result) === view.filter) : view.items;
  if (!items.length) {
    area.innerHTML = `<div class="empty"><p class="empty__title">該当する結果がありません</p>
      <p class="empty__text">絞り込みを解除すると、すべての結果を表示します。</p></div>`;
    return;
  }

  area.innerHTML = `<table class="table">
    <colgroup>
      <col class="col-result"><col class="col-name"><col class="col-month"><col class="col-rule">
      <col class="col-planned"><col class="col-actual"><col class="col-message">
    </colgroup>
    <thead><tr>
      <th scope="col">結果</th>
      <th scope="col">ファイル名</th>
      <th scope="col">抽出年月</th>
      <th scope="col">判定ルール</th>
      <th scope="col">通常保存予定先</th>
      <th scope="col">実際の保存先${view.mode === 'preview' ? '予定' : ''}</th>
      <th scope="col">メッセージ</th>
    </tr></thead>
    <tbody>${items.map((item, index) => resultRow(item, index, view.mode)).join('')}</tbody>
  </table>`;
}

function resultRow(item, index, mode) {
  const key = `${item.file_name}#${index}`;
  const open = state.expanded.has(key);
  const actual = item.actual_destination || item.destination || '';
  const planned = item.planned_destination || '';

  const main = `<tr data-key="${esc(key)}">
    <td>${resultChip(item.result, mode)}</td>
    <td class="cell-name"><button type="button" class="row-toggle" data-action="toggle-detail" title="クリックで根拠と全文パスを表示">
      <span class="truncate" title="${esc(item.file_name)}">${esc(item.file_name)}</span></button></td>
    <td class="cell-month">${esc(item.extracted_month || '—')}</td>
    <td><span class="truncate" title="${esc(item.rule_name || '一致なし')}">${esc(item.rule_name || '—')}</span></td>
    <td class="cell-path"><span class="truncate" title="${esc(planned)}">${esc(shortenPath(folderOf(planned)) || '—')}</span></td>
    <td class="cell-path"><span class="truncate" title="${esc(actual)}">${esc(shortenPath(folderOf(actual)) || '—')}</span></td>
    <td class="cell-message"><span class="truncate" title="${esc(item.message || '')}">${esc(item.message || '')}</span></td>
  </tr>`;

  if (!open) return main;

  const reasons = (item.reason || []).map((reason) => `<span>・${esc(reason)}</span>`).join('');
  return main + `<tr class="detail"><td colspan="7">
    <dl>
      <dt>結果コード</dt><dd>${esc(item.result)}</dd>
      <dt>移動元</dt><dd>${esc(item.source)}</dd>
      <dt>通常保存予定先</dt><dd>${esc(planned || '（なし）')}</dd>
      <dt>実際の保存先</dt><dd>${esc(actual || '（移動しません）')}</dd>
      <dt>判定の根拠</dt><dd class="reasons">${reasons || '—'}</dd>
    </dl>
  </td></tr>`;
}

function fileList(files) {
  return `<div class="filelist">${files.map((file) => {
    const blocked = file.blocked;
    return `<div class="filelist__row">
      <span class="filelist__name" title="${esc(file.name)}">${esc(file.name)}</span>
      <span class="filelist__meta">${esc(formatBytes(file.size))} ／ 更新 ${esc(file.modified)}</span>
      ${blocked
        ? `<span class="chip chip--neutral" title="${esc(blocked)}"><span class="chip__icon" aria-hidden="true">—</span>スキップ予定</span>`
        : `<span class="chip chip--info"><span class="chip__icon" aria-hidden="true">→</span>処理できます</span>`}
    </div>${blocked ? `<p class="field__hint" style="padding-left:var(--space-3)">${esc(blocked)}</p>` : ''}`;
  }).join('')}</div>`;
}

function renderActionBar() {
  const meta = state.meta;
  const count = meta ? meta.file_count : 0;
  const validation = validate();
  const base = baseError();
  const sortButton = $('#btn-sort');
  const reason = $('#cta-reason');
  const hint = $('#cta-hint');

  $('#btn-save').disabled = state.busy || validation.messages.length > 0;
  $('#btn-preview').disabled = state.busy || validation.messages.length > 0;
  $('#btn-reload').disabled = state.busy;

  let label = `${count} 件のPDFを移動する`;
  let disabled = false;
  let reasonHtml = '';
  let hintText = '';

  if (state.busy) {
    label = '処理しています…';
    disabled = true;
    reasonHtml = '処理が終わるまでお待ちください';
  } else if (state.fatal) {
    disabled = true;
    reasonHtml = `<span class="chip chip--danger"><span class="chip__icon" aria-hidden="true">✕</span>実行できません</span>rules.json を読み込めません`;
    hintText = meta ? `${meta.rules_path} を修正するか、削除して再読み込みしてください（削除時は初期設定で自動作成されます）` : '';
  } else if (validation.messages.length) {
    disabled = true;
    reasonHtml = `<span class="chip chip--danger"><span class="chip__icon" aria-hidden="true">✕</span>実行できません</span>設定に ${validation.messages.length} 件の問題があります`;
    hintText = validation.messages[0] + (validation.messages.length > 1 ? ` ほか ${validation.messages.length - 1} 件` : '');
  } else if (base) {
    disabled = true;
    reasonHtml = `<span class="chip chip--danger"><span class="chip__icon" aria-hidden="true">✕</span>実行できません</span>${esc(base)}`;
    hintText = '「共通設定」で共通保存先フォルダを指定してください';
  } else if (meta && meta.lock && meta.lock.locked) {
    disabled = true;
    reasonHtml = `<span class="chip chip--warn"><span class="chip__icon" aria-hidden="true">▲</span>実行中</span>他の処理が実行中です`;
    hintText = `ロックファイル: ${meta.lock.path}（実行中のウィンドウが無い場合は削除してください）`;
  } else if (count === 0) {
    label = '移動するPDFがありません';
    disabled = true;
    reasonHtml = `<span class="chip chip--neutral"><span class="chip__icon" aria-hidden="true">—</span>対象なし</span>このフォルダに対象のPDFがありません`;
    hintText = meta ? `${meta.script_folder} にPDFを置いてから「再読み込み」を押してください` : '';
  } else {
    const enabledRules = state.rules.filter((rule) => rule.enabled).length;
    const bases = new Set(state.rules.filter((rule) => rule.enabled).map((rule) => ruleBase(rule).path));
    const target = bases.size === 1
      ? `${esc(tailText([...bases][0], 40))} 配下へ移動します`
      : `${bases.size} か所の保存先へ振り分けます`;
    reasonHtml = enabledRules
      ? `<span class="chip chip--info"><span class="chip__icon" aria-hidden="true">→</span>実行できます</span>${count} 件を判定して ${target}`
      : `<span class="chip chip--warn"><span class="chip__icon" aria-hidden="true">▲</span>ルールなし</span>有効なルールが 0 件です。すべて判定不能フォルダへ移動します`;
    hintText = state.dirty
      ? '未保存の変更があります（実行すると rules.json も保存されます）'
      : '実行前に rules.json を保存します。移動後は取り消しできます。';
  }

  sortButton.textContent = label;
  sortButton.disabled = disabled;
  reason.innerHTML = reasonHtml;
  reason.title = reason.textContent;
  hint.textContent = hintText;
  hint.title = hintText;
}

/* =======================================================================
   共通設定ダイアログ
   ===================================================================== */

const MONTH_FALLBACK_HINTS = {
  unknown_month: '年月が取れないPDFは _年月不明 フォルダへ退避します（実行月へ誤って保存しないため推奨）',
  current_month: '年月が取れないPDFは実行日の年月フォルダへ保存します',
  error: '年月が取れないPDFは移動せず、エラーとして記録します',
};

const DUPLICATE_HINTS = {
  evacuate: '保存先に同名ファイルがある場合、年月フォルダ内の 避難用_YYYYMMDD へ保存します（上書きしません）',
  skip: '保存先に同名ファイルがある場合、そのPDFは移動しません',
};

const MOVE_STRATEGY_HINTS = {
  copy_verify_delete: 'コピー後に存在とサイズを確認し、確認できてから元ファイルを削除します（紛失リスクが最も低い方式）',
  move: 'コピー後に元ファイルを削除します。サイズ検証は行いません',
  copy_only: 'コピーだけを行い、元ファイルはこのフォルダに残ります',
};

function bindSettingsDialog() {
  document.querySelectorAll('[data-setting]').forEach((input) => {
    const key = input.dataset.setting;
    let value = state.settings[key];
    if (key === 'common_base' && value === '{BASE_FOLDER}') value = '';
    if (input.type === 'checkbox') input.checked = !!value;
    else input.value = value === undefined || value === null ? '' : value;
  });
  updateSettingHints();
}

function updateSettingHints() {
  const meta = state.meta || {};
  const base = String(state.settings.common_base || '').trim();
  const problem = baseError();
  const firstRule = state.rules.find((rule) => rule.destination_subfolder);
  const sample = joinPath(base || '{共通保存先}', firstRule ? firstRule.destination_subfolder : '{保存先サブフォルダ}',
    state.settings.use_month_folder ? '2026-08' : '');

  $('#settings-lead').hidden = !problem;
  $('#hint-common-base').textContent = problem
    ? `${problem}（例: D:\\共有\\設備記録 のように絶対パスで指定します）`
    : commonBaseRequired()
      ? `保存先フォルダを指定していないルールが使います。保存先の例: ${tailText(sample, 48)}`
      : 'いまは使われていません（すべての有効なルールが専用の保存先フォルダを指定しています）';
  const unknownPath = joinPath(meta.script_folder || '', state.settings.unknown_folder || '');
  const logFilePath = joinPath(meta.script_folder || '', state.settings.log_folder || '', 'move_log.csv');
  $('#hint-unknown-folder').textContent = `ルールに一致しないPDFの退避先: ${tailText(unknownPath, 40)}`;
  $('#hint-unknown-folder').title = unknownPath;
  $('#hint-log-folder').textContent = `移動結果CSVの保存先: ${tailText(logFilePath, 40)}`;
  $('#hint-log-folder').title = logFilePath;
  $('#hint-use-month').textContent = state.settings.use_month_folder
    ? '保存先サブフォルダの下に YYYY-MM フォルダを作成します'
    : '年月フォルダを作らず、保存先サブフォルダへ直接保存します';
  $('#hint-month-fallback').textContent = MONTH_FALLBACK_HINTS[state.settings.month_fallback] || '';
  $('#hint-duplicate-mode').textContent = DUPLICATE_HINTS[state.settings.duplicate_mode] || '';
  $('#hint-move-strategy').textContent = MOVE_STRATEGY_HINTS[state.settings.move_strategy] || '';
  $('#settings-foot-note').textContent = state.dirty
    ? '変更は画面上に保持されています。保存するには「rules.json に保存」を押してください。'
    : '変更内容は保存または実行のときに rules.json へ書き込まれます。';

  const patterns = (meta.month_patterns || []).join(' / ');
  $('#fixed-settings').innerHTML = [
    ['既存ファイルの上書き', '常に禁止（事故防止の必須要件のため変更できません）'],
    ['年月フォルダの形式', `${state.settings.month_folder_format}（初期仕様で固定）`],
    ['年月の取得元', 'ファイル名（初期仕様で固定）'],
    ['認識する年月表記', patterns],
    ['年月が不明なとき', `${meta.month_unknown_folder || '_年月不明'} へ退避`],
    ['年月候補が複数のとき', `${meta.month_ambiguous_folder || '_年月確認要'} へ退避`],
    ['同名時の避難先', `${meta.evacuation_example || '避難用_YYYYMMDD'}（連番は避難先で同名の場合のみ）`],
    ['rules.json のバックアップ', `保存のたびに作成: ${meta.backup_folder || ''}`],
    ['通信範囲', '127.0.0.1 のみ。外部へは送信しません'],
  ].map(([term, description]) => `<li><b>${esc(term)}</b><span>${esc(description)}</span></li>`).join('');
}

/* =======================================================================
   操作
   ===================================================================== */

function markDirty() {
  state.dirty = true;
  renderActionBar();
}

function toast(message, actions) {
  const list = actions ? (Array.isArray(actions) ? actions : [actions]) : [];
  const box = $('#toasts');
  const element = document.createElement('div');
  element.className = 'toast';
  element.innerHTML = `<span>${esc(message)}</span>`;
  list.forEach((action) => {
    const button = document.createElement('button');
    button.type = 'button';
    button.className = 'btn btn--quiet';
    button.textContent = action.label;
    button.addEventListener('click', () => {
      element.remove();
      action.run();
    });
    element.appendChild(button);
  });
  box.appendChild(element);
  setTimeout(() => element.remove(), list.length ? 12000 : 5000);
}

function nextPriority() {
  return state.rules.reduce((max, rule) => Math.max(max, Number(rule.priority) || 0), 0) + 1;
}

function addRule() {
  state.rules.push({
    _id: state.nextId++,
    enabled: true,
    priority: nextPriority(),
    name: '',
    extension: '.pdf',
    contains_all_text: '',
    contains_any_text: '',
    not_contains_text: '',
    destination_base: '',
    destination_subfolder: '',
  });
  markDirty();
  renderRules();
  renderActionBar();
  const inputs = document.querySelectorAll('.rule__name');
  if (inputs.length) inputs[inputs.length - 1].focus();
}

function deleteRule(id) {
  const index = state.rules.findIndex((rule) => rule._id === id);
  if (index < 0) return;
  const [removed] = state.rules.splice(index, 1);
  markDirty();
  renderRules();
  renderActionBar();
  toast(`ルール「${removed.name || '名称未設定'}」を削除しました`, {
    label: '取り消す',
    run: () => {
      state.rules.splice(index, 0, removed);
      renderRules();
      renderActionBar();
    },
  });
}

function clearRules() {
  if (!state.rules.length) { toast('削除できるルールがありません'); return; }
  const backup = state.rules.slice();
  state.rules = [];
  markDirty();
  renderRules();
  renderActionBar();
  toast(`${backup.length} 件のルールをすべて削除しました`, {
    label: '取り消す',
    run: () => {
      state.rules = backup;
      renderRules();
      renderActionBar();
    },
  });
}

function moveRule(id, direction) {
  const index = state.rules.findIndex((rule) => rule._id === id);
  const target = index + direction;
  if (index < 0 || target < 0 || target >= state.rules.length) return;
  const rules = state.rules;
  [rules[index], rules[target]] = [rules[target], rules[index]];
  rules.forEach((rule, position) => { rule.priority = position + 1; });
  markDirty();
  renderRules();
}

async function doSave(silent) {
  if (state.busy) return false;
  state.busy = true;
  renderActionBar();
  const { data, status } = await api('/api/rules', 'POST', currentPayload());
  state.busy = false;

  if (!data.ok) {
    renderActionBar();
    toast(data.message || `保存できませんでした（HTTP ${status}）`);
    (data.errors || []).slice(0, 3).forEach((error) => toast(error));
    return false;
  }
  state.dirty = false;
  adoptMeta(data._meta);
  renderFacts();
  renderActionBar();
  renderUndoBar();
  if (!silent) {
    toast(data.backup_path ? 'rules.json を保存しました（バックアップも作成しました）' : 'rules.json を保存しました');
  }
  return true;
}

async function doPreview() {
  if (state.busy) return;
  state.busy = true;
  renderActionBar();
  const { data, status } = await api('/api/preview', 'POST', currentPayload());
  state.busy = false;

  if (data.items === undefined) {
    renderActionBar();
    toast(data.message || `プレビューできませんでした（HTTP ${status}）`);
    (data.errors || []).slice(0, 3).forEach((error) => toast(error));
    return;
  }
  adoptMeta(data._meta);
  state.expanded.clear();
  state.view = {
    mode: 'preview', items: data.items, summary: data.summary,
    at: new Date().toLocaleTimeString('ja-JP', { hour: '2-digit', minute: '2-digit' }), filter: null,
  };
  renderAll();
  if (data.base_folder_error) toast(data.base_folder_error);
}

async function doSort() {
  if (state.busy) return;
  state.busy = true;
  renderActionBar();
  const { data, status } = await api('/api/sort', 'POST', currentPayload());
  state.busy = false;

  if (data.items === undefined) {
    renderActionBar();
    toast(data.message || `実行できませんでした（HTTP ${status}）`);
    (data.errors || []).slice(0, 3).forEach((error) => toast(error));
    return;
  }
  state.dirty = false;
  adoptMeta(data._meta);
  state.expanded.clear();
  state.view = {
    mode: 'sorted', items: data.items, summary: data.summary,
    at: new Date().toLocaleTimeString('ja-JP', { hour: '2-digit', minute: '2-digit' }), filter: null,
  };
  renderAll();
  toast(data.message || '移動しました');
}

async function doUndo() {
  if (state.busy) return;
  state.busy = true;
  renderUndoBar();
  renderActionBar();
  const { data, status } = await api('/api/undo', 'POST', {});
  state.busy = false;

  if (data.items === undefined) {
    renderActionBar();
    toast(data.message || `取り消せませんでした（HTTP ${status}）`);
    return;
  }
  adoptMeta(data._meta);
  state.expanded.clear();
  state.view = {
    mode: 'undone', items: data.items, summary: summarizeItems(data.items),
    at: new Date().toLocaleTimeString('ja-JP', { hour: '2-digit', minute: '2-digit' }), filter: null,
  };
  renderAll();
  toast(data.message || '取り消しました');
}

function summarizeItems(items) {
  const summary = { success: 0, evacuated: 0, month: 0, unknown: 0, error: 0, skip: 0, total: items.length };
  items.forEach((item) => { summary[categorize(item.result)] += 1; });
  return summary;
}

/* ツールの終了（コンソールを触らずに済ませる） */
function quitTool() {
  if (state.dirty) {
    toast('未保存の変更があります。どうしますか？', [
      { label: '保存して終了', run: async () => { if (await doSave(true)) shutdown(); } },
      { label: '保存せず終了', run: shutdown },
    ]);
    return;
  }
  shutdown();
}

async function shutdown() {
  state.dirty = false;
  await api('/api/shutdown', 'POST', {});
  state.closed = true;
  $('#farewell').hidden = false;
}

async function copyText(text, label) {
  try {
    await navigator.clipboard.writeText(text);
    toast(`${label}をコピーしました`);
  } catch {
    toast(`コピーできませんでした: ${text}`);
  }
}

/* =======================================================================
   イベント
   ===================================================================== */

function bindEvents() {
  $('#btn-add-rule').addEventListener('click', addRule);
  $('#btn-reload').addEventListener('click', () => loadAll(false));
  $('#btn-save').addEventListener('click', () => doSave(false));
  $('#btn-preview').addEventListener('click', doPreview);
  $('#btn-sort').addEventListener('click', doSort);

  $('#btn-settings').addEventListener('click', () => {
    bindSettingsDialog();
    $('#settings-dialog').showModal();
  });

  // ルール一覧（入力とボタンは委譲で処理する）
  const list = $('#rule-list');
  list.addEventListener('input', (event) => {
    const card = event.target.closest('.rule');
    const fieldName = event.target.dataset.field;
    if (!card || !fieldName || fieldName === 'priority') return;
    const rule = state.rules.find((entry) => entry._id === Number(card.dataset.id));
    if (!rule) return;
    rule[fieldName] = event.target.type === 'checkbox' ? event.target.checked : event.target.value;
    markDirty();
    updateCardFoot(card, rule);
  });

  list.addEventListener('change', (event) => {
    const card = event.target.closest('.rule');
    const fieldName = event.target.dataset.field;
    if (!card || !fieldName) return;
    const rule = state.rules.find((entry) => entry._id === Number(card.dataset.id));
    if (!rule) return;
    if (fieldName === 'enabled') {
      rule.enabled = event.target.checked;
      markDirty();
      renderRules();
    } else if (fieldName === 'priority') {
      rule.priority = Number(event.target.value) || 0;
      markDirty();
      sortRules();
      renderRules();
    } else {
      renderRules();
    }
  });

  list.addEventListener('click', (event) => {
    const button = event.target.closest('button[data-action]');
    if (!button) return;
    const card = button.closest('.rule');
    const id = Number(card.dataset.id);
    const action = button.dataset.action;
    if (action === 'delete') deleteRule(id);
    else if (action === 'up') moveRule(id, -1);
    else if (action === 'down') moveRule(id, 1);
    else if (action === 'duplicate') {
      const source = state.rules.find((rule) => rule._id === id);
      const copy = Object.assign({}, source, { _id: state.nextId++, priority: nextPriority(), name: `${source.name} のコピー` });
      state.rules.push(copy);
      markDirty();
      renderRules();
    }
  });

  // 結果パネル
  $('#result-area').addEventListener('click', (event) => {
    const toggle = event.target.closest('[data-action="toggle-detail"]');
    if (!toggle) return;
    const key = toggle.closest('tr').dataset.key;
    if (state.expanded.has(key)) state.expanded.delete(key);
    else state.expanded.add(key);
    renderResults();
  });

  $('#summary').addEventListener('click', (event) => {
    const button = event.target.closest('[data-filter]');
    if (!button) return;
    const value = button.dataset.filter;
    state.view.filter = value === 'clear' || state.view.filter === value ? null : value;
    renderSummary();
    renderResults();
  });

  $('#result-head-actions').addEventListener('click', (event) => {
    if (!event.target.closest('[data-action="back-to-files"]')) return;
    state.view = { mode: 'idle', items: [], summary: null, at: '', filter: null };
    renderSummary();
    renderResults();
  });

  $('#undo-bar').addEventListener('click', (event) => {
    if (event.target.id === 'btn-undo') doUndo();
  });

  // 共通設定
  document.querySelectorAll('[data-setting]').forEach((input) => {
    const handler = () => {
      const key = input.dataset.setting;
      if (input.type === 'checkbox') state.settings[key] = input.checked;
      else if (input.type === 'number') state.settings[key] = Number(input.value);
      else state.settings[key] = input.value;
      markDirty();
      updateSettingHints();
      renderFacts();
      renderRules();
    };
    input.addEventListener('input', handler);
    input.addEventListener('change', handler);
  });

  // オーバーフローメニュー（危ない操作はここに置く）
  const menuButton = $('#btn-menu');
  const menuList = $('#menu-list');
  menuButton.addEventListener('click', () => {
    const open = menuList.hidden;
    menuList.hidden = !open;
    menuButton.setAttribute('aria-expanded', String(open));
  });
  document.addEventListener('click', (event) => {
    if (!event.target.closest('.menu')) {
      menuList.hidden = true;
      menuButton.setAttribute('aria-expanded', 'false');
    }
  });
  menuList.addEventListener('click', (event) => {
    const button = event.target.closest('[data-menu]');
    if (!button) return;
    menuList.hidden = true;
    menuButton.setAttribute('aria-expanded', 'false');
    const action = button.dataset.menu;
    if (action === 'reload') loadAll(false);
    else if (action === 'clear-rules') clearRules();
    else if (action === 'copy-log') copyText(state.meta ? state.meta.log_path : '', 'ログのパス');
    else if (action === 'copy-backup') copyText(state.meta ? state.meta.backup_folder : '', 'バックアップ先のパス');
    else if (action === 'quit') quitTool();
  });

  // キーボード
  document.addEventListener('keydown', (event) => {
    if ((event.ctrlKey || event.metaKey) && event.key === 's') {
      event.preventDefault();
      if (!$('#btn-save').disabled) doSave(false);
    }
  });

  // 未保存のまま閉じるのを防ぐ
  window.addEventListener('beforeunload', (event) => {
    if (!state.dirty) return;
    event.preventDefault();
    event.returnValue = '';
  });

  // 画面に戻ったときはフォルダの状況を取り直す
  window.addEventListener('focus', () => { if (!state.busy) refreshMeta(); });
}

function updateCardFoot(card, rule) {
  const code = card.querySelector('.rule__foot code');
  if (!code) return;
  code.textContent = destinationExample(rule, false);
  code.title = destinationExample(rule, true);
}

/* =======================================================================
   起動
   ===================================================================== */

bindEvents();
loadAll(true);
"""


if __name__ == "__main__":
    sys.exit(main(sys.argv))
