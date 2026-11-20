# -*- coding: utf-8 -*-
"""装箱配置 (PackProfile) — 与 UI 无关，可被其他程序 import。

接口与 MX 包装配置 (VariationProfile) 对齐：
    get(id) / all_profiles() / profile_from_dict(d) / profile_to_dict(p)
自定义配置保存在程序配置 JSON 的 ``custom_pack_profiles`` 键下
（profile_to_dict 的列表），不同程序读写同一文件即可共享。
另外会尝试只读导入 MX 工具 ``%APPDATA%\\装箱单工具\\config.json``
中的 ``mx_custom_variations``。
"""

import json
import os
import re
from dataclasses import dataclass, field, replace, asdict   # noqa: F401 (replace re-exported)

__all__ = [
    'PackProfile', 'BUILTINS', 'get', 'all_profiles',
    'profile_from_dict', 'profile_to_dict',
    'load_custom_profiles', 'dump_custom_profiles',
    'cap_for_size', 'parse_tiers_text', 'tiers_to_text', 'replace',
]


@dataclass(frozen=True)
class PackProfile:
    id: str
    name_zh: str
    method: str = 'loose'            # loose | ratio | set（本引擎只实现 loose）
    sizes: tuple = ()                # 有序尺码表；空=按数据自动排序
    tiers: tuple = ()                # ((起始尺码或'*', 每箱数量), ...)
    net_k: float = 0.0               # 建议单件净重 kg（生成时可覆盖）
    gross_add: float = 0.8           # 大箱纸箱重量 kg
    gross_add_mid: float = 0.6       # 中箱
    gross_add_small: float = 0.5     # 小箱
    allowance: int = 3               # 尾箱/混码箱允许超出件数
    pieces_per_pack: int = 1         # 1=按件, 2/3/4=按套
    cap_mid: int = 0                 # 中箱默认容量（件）；0=未设置
    cap_small: int = 0               # 小箱默认容量（件）；0=未设置

    @property
    def unit(self) -> str:
        return '件' if self.pieces_per_pack <= 1 else '套'


# ── 内置配置（来自 2026 Q2 再生系列装箱单分析）──────────────────
BUILTINS = [
    PackProfile(id='标准90',  name_zh='标准·90件/箱（文胸类）',
                tiers=(('*', 90),),  net_k=0.063),
    PackProfile(id='标准260', name_zh='标准·260件/箱',
                tiers=(('*', 260),), net_k=0.046),
    PackProfile(id='标准300', name_zh='标准·300件/箱（内裤类）',
                tiers=(('*', 300),), net_k=0.044),
    PackProfile(id='套装2件', name_zh='套装·2件套（按套装箱）',
                tiers=(('*', 90),),  pieces_per_pack=2),
]
_BUILTIN_BY_ID = {p.id: p for p in BUILTINS}

# MX 工具的共享配置文件（只读导入）
MX_CONFIG_PATH = os.path.expandvars(r'%APPDATA%\装箱单工具\config.json')


def profile_to_dict(p: PackProfile) -> dict:
    d = asdict(p)
    d['sizes'] = list(p.sizes)
    d['tiers'] = [list(t) for t in p.tiers]
    return d


def profile_from_dict(d: dict) -> PackProfile:
    pid = str(d.get('id') or '').strip()
    if not pid:
        raise ValueError('配置缺少 id')
    if pid not in _BUILTIN_BY_ID and not pid.startswith('custom:'):
        pid = f'custom:{pid}'
    tiers = tuple((str(s), int(n)) for s, n in (d.get('tiers') or []))
    # carton_big/mid/small keys ignored — dims are now auto-detected by port
    return PackProfile(
        id=pid,
        name_zh=str(d.get('name_zh') or pid),
        method=str(d.get('method') or 'loose'),
        sizes=tuple(d.get('sizes') or ()),
        tiers=tiers,
        net_k=float(d.get('net_k') or 0),
        gross_add=float(d.get('gross_add') or 0.8),
        gross_add_mid=float(d.get('gross_add_mid') or 0.6),
        gross_add_small=float(d.get('gross_add_small') or 0.5),
        allowance=int(d.get('allowance') or 0),
        pieces_per_pack=int(d.get('pieces_per_pack') or 1),
        cap_mid=int(d.get('cap_mid') or 0),
        cap_small=int(d.get('cap_small') or 0),
    )


def load_custom_profiles(cfg: dict) -> list:
    """从本程序配置 dict 读取自定义配置 + 只读导入 MX 共享配置。"""
    profiles = []
    for d in cfg.get('custom_pack_profiles', []):
        try:
            profiles.append(profile_from_dict(d))
        except Exception:
            continue
    # MX 共享（只读；同 id 不重复）
    try:
        with open(MX_CONFIG_PATH, encoding='utf-8') as f:
            mx_cfg = json.load(f)
        have = {p.id for p in profiles} | set(_BUILTIN_BY_ID)
        for d in mx_cfg.get('mx_custom_variations', []):
            try:
                p = profile_from_dict(d)
            except Exception:
                continue
            if p.id not in have:
                profiles.append(p)
                have.add(p.id)
    except Exception:
        pass
    return profiles


def dump_custom_profiles(profiles) -> list:
    """仅导出本程序自定义配置（custom: 前缀）为 JSON-able 列表。"""
    return [profile_to_dict(p) for p in profiles if p.id.startswith('custom:')]


def all_profiles(cfg: dict = None) -> list:
    customs = load_custom_profiles(cfg or {}) if cfg is not None else []
    custom_names = {p.name_zh for p in customs}
    # Built-ins whose name_zh has a custom override are hidden so the custom takes effect.
    return [b for b in BUILTINS if b.name_zh not in custom_names] + customs


def get(pid: str, cfg: dict = None) -> PackProfile:
    for p in all_profiles(cfg):
        if p.id == pid or p.name_zh == pid:
            return p
    raise KeyError(f'未找到包装配置: {pid}')


# ── 分级容量 ──────────────────────────────────────────────────

def cap_for_size(profile: PackProfile, size: str, size_order=None) -> int:
    """按 tiers 查该尺码的每箱数量。

    tiers 形如 (('*', 90),) 或 (('2XS', 90), ('L', 80))：
    从起始尺码（含）到下一级起始尺码（不含）适用该级容量。
    尺码顺序优先用 profile.sizes，否则用 size_order（数据中的有序尺码表）。
    """
    if not profile.tiers:
        raise ValueError(f'配置 {profile.id} 没有 tiers')
    order = list(profile.sizes) or list(size_order or [])
    cap = profile.tiers[0][1]
    if not order:
        return cap
    try:
        pos = order.index(size)
    except ValueError:
        return cap
    for start, n in profile.tiers:
        if start == '*':
            cap = n
            continue
        try:
            start_pos = order.index(start)
        except ValueError:
            continue
        if pos >= start_pos:
            cap = n
    return cap


# ── tiers 文本互转（UI 编辑用）─────────────────────────────────

def parse_tiers_text(text: str) -> tuple:
    """'90' → (('*',90),)；'2XS:90, L:80' → (('2XS',90),('L',80))"""
    text = str(text or '').strip()
    if not text:
        return ()
    parts = [p.strip() for p in re.split(r'[,，;；]', text) if p.strip()]
    tiers = []
    for p in parts:
        if ':' in p or '：' in p:
            s, n = re.split(r'[:：]', p, 1)
            tiers.append((s.strip(), int(float(n))))
        else:
            tiers.append(('*', int(float(p))))
    return tuple(tiers)


def tiers_to_text(tiers) -> str:
    return ', '.join(str(n) if s == '*' else f'{s}:{n}' for s, n in tiers)
