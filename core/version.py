"""core/version.py — 版本号的唯一真源。

README / CITATION.cff / 界面 / run.py 都从这里取，改版本只改一行。
（发布流程：改这里 → 同步英文 README 版本表 → git tag -a vX.Y.Z → 发 Release。）
"""
from __future__ import annotations

__version__ = "3.0.0"
