"""Reusable fill templates stored in the user's config directory.

A template is a small JSON document describing everything needed to repeat a
fill run: target workbook, column mapping, extraction rules, output settings…
Templates live outside the project directory so they survive upgrades and are
shared by every session.

Layout::

    <config>/TableFiller/templates/<kind>/<name>.json

``kind`` is one of ``image``, ``text`` or ``cross``.
"""

from __future__ import annotations

import datetime as _dt
import json
import os
import re
import sys
from dataclasses import dataclass, field

from .errors import DataError
from .safeio import write_json

SCHEMA_VERSION = 1
KINDS = ("image", "text", "cross")
KIND_LABELS = {
    "image": "图片识别填表",
    "text": "文字填表",
    "cross": "跨文件填表",
}

_INVALID_NAME = re.compile(r'[\\/:*?"<>|\x00-\x1f]')


def templates_root() -> str:
    """Return the platform-specific directory that stores templates."""
    override = os.environ.get("TABLE_FILLER_HOME")
    if override:
        base = os.path.abspath(os.path.expanduser(override))
    elif os.name == "nt":
        base = os.environ.get("APPDATA") or os.path.expanduser("~")
    elif sys.platform == "darwin":
        base = os.path.expanduser("~/Library/Application Support")
    else:
        base = os.environ.get("XDG_CONFIG_HOME") or os.path.expanduser("~/.config")
    return os.path.join(base, "TableFiller", "templates")


def validate_name(name: str) -> str:
    """Reject names that could escape the template directory."""
    name = (name or "").strip()
    if not name:
        raise DataError("模板名称不能为空。")
    if len(name) > 80:
        raise DataError("模板名称不能超过 80 个字符。")
    if name in (".", "..") or _INVALID_NAME.search(name):
        raise DataError("模板名称不能包含路径分隔符或 \\ / : * ? \" < > | 等字符。")
    return name


def _check_kind(kind: str) -> str:
    if kind not in KINDS:
        raise DataError(f"未知的模板类型：{kind}")
    return kind


def _kind_dir(kind: str) -> str:
    return os.path.join(templates_root(), _check_kind(kind))


def _template_path(kind: str, name: str) -> str:
    return os.path.join(_kind_dir(kind), validate_name(name) + ".json")


def _now() -> str:
    return _dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")


@dataclass
class Template:
    name: str
    kind: str
    data: dict = field(default_factory=dict)
    created: str = ""
    updated: str = ""
    version: int = SCHEMA_VERSION
    path: str = ""

    def to_dict(self) -> dict:
        return {
            "schema": self.version,
            "name": self.name,
            "kind": self.kind,
            "created": self.created,
            "updated": self.updated,
            "data": self.data,
        }

    @classmethod
    def from_dict(cls, payload: dict, path: str = "") -> "Template":
        if not isinstance(payload, dict):
            raise DataError("模板文件格式不正确。")
        version = int(payload.get("schema", 0) or 0)
        if version > SCHEMA_VERSION:
            raise DataError(
                f"模板版本 {version} 高于当前程序支持的 {SCHEMA_VERSION}，请升级程序。")
        kind = payload.get("kind", "")
        if kind not in KINDS:
            raise DataError(f"模板文件中的类型无效：{kind}")
        name = payload.get("name") or (os.path.splitext(os.path.basename(path))[0] if path else "")
        data = payload.get("data")
        if not isinstance(data, dict):
            raise DataError("模板内容缺失或格式不正确。")
        return cls(
            name=name,
            kind=kind,
            data=data,
            created=payload.get("created", ""),
            updated=payload.get("updated", ""),
            version=version or SCHEMA_VERSION,
            path=path,
        )

    def summary(self) -> str:
        return f"{KIND_LABELS.get(self.kind, self.kind)} · {self.name}（更新于 {self.updated}）"


def _read(path: str) -> Template:
    try:
        with open(path, "r", encoding="utf-8-sig") as fh:
            payload = json.load(fh)
    except json.JSONDecodeError as exc:
        raise DataError(f"模板文件解析失败：{exc}") from exc
    except OSError as exc:
        raise DataError(f"无法读取模板文件：{exc}") from exc
    return Template.from_dict(payload, path)


def save_template(name: str, kind: str, data: dict, overwrite: bool = True) -> Template:
    """Create or update a template and return it."""
    name = validate_name(name)
    _check_kind(kind)
    if not isinstance(data, dict):
        raise DataError("模板内容必须是对象。")
    path = _template_path(kind, name)
    created = _now()
    version = SCHEMA_VERSION
    if os.path.exists(path):
        if not overwrite:
            raise DataError(f"模板已存在：{name}")
        try:
            existing = _read(path)
            created = existing.created or created
            version = existing.version or version
        except DataError:
            pass
    payload = {
        "schema": version,
        "name": name,
        "kind": kind,
        "created": created,
        "updated": _now(),
        "data": data,
    }
    try:
        write_json(path, payload)
    except OSError as exc:
        raise DataError(f"无法保存模板：{exc}") from exc
    return Template(name=name, kind=kind, data=data, created=created,
                    updated=payload["updated"], version=version, path=path)


def load_template(name: str, kind: str = "") -> Template:
    """Load one template, searching every kind when *kind* is empty."""
    name = validate_name(name)
    kinds = [_check_kind(kind)] if kind else list(KINDS)
    for candidate in kinds:
        path = _template_path(candidate, name)
        if os.path.isfile(path):
            return _read(path)
    raise DataError(f"找不到模板：{name}")


def list_templates(kind: str = "") -> list[Template]:
    """List stored templates, newest first."""
    kinds = [_check_kind(kind)] if kind else list(KINDS)
    found: list[Template] = []
    for candidate in kinds:
        directory = _kind_dir(candidate)
        if not os.path.isdir(directory):
            continue
        for entry in os.listdir(directory):
            if not entry.lower().endswith(".json"):
                continue
            path = os.path.join(directory, entry)
            try:
                found.append(_read(path))
            except DataError:
                continue
    found.sort(key=lambda item: item.updated, reverse=True)
    return found


def delete_template(name: str, kind: str) -> bool:
    path = _template_path(kind, name)
    if not os.path.isfile(path):
        return False
    try:
        os.remove(path)
    except OSError as exc:
        raise DataError(f"无法删除模板：{exc}") from exc
    return True


def export_template(name: str, kind: str, destination: str) -> str:
    """Copy a template JSON to *destination* (e.g. to share with a colleague)."""
    template = load_template(name, kind)
    destination = os.path.abspath(destination)
    if not destination.lower().endswith(".json"):
        destination += ".json"
    directory = os.path.dirname(destination)
    if directory:
        os.makedirs(directory, exist_ok=True)
    try:
        write_json(destination, template.to_dict())
    except OSError as exc:
        raise DataError(f"无法导出模板：{exc}") from exc
    return destination


def import_template(source: str, kind: str = "", name: str = "") -> Template:
    """Import a template JSON file and store it under the given/its own name."""
    if not os.path.isfile(source):
        raise DataError(f"模板文件不存在：{source}")
    template = _read(source)
    if kind:
        template.kind = _check_kind(kind)
    template.name = validate_name(name or template.name)
    return save_template(template.name, template.kind, template.data, overwrite=True)
