#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
引擎模塊

提供統一的引擎選擇和載入接口

作者: David-CB666
版本: v2.1
日期: 2026-06-02
"""

from .base_engine import BaseEngine
from .openpyxl_engine import OpenPYXLEngine
from .zip_engine import ZIPEngine


def create_engine(template_path: str = None, engine_type: str = "auto") -> BaseEngine:
    """
    創建引擎實例
    
    Args:
        template_path: 模板路徑（用於自動檢測）
        engine_type: 引擎類型
            "auto" - 自動檢測
            "openpyxl" - 使用 openpyxl 引擎
            "zip" - 使用 ZIP 引擎
    
    Returns:
        引擎實例
    """
    if engine_type == "auto" and template_path:
        # 自動檢測
        engine_type = BaseEngine.detect_template_type(template_path)
    
    if engine_type == "zip":
        return ZIPEngine()
    else:
        return OpenPYXLEngine()


__all__ = [
    "BaseEngine",
    "OpenPYXLEngine",
    "ZIPEngine",
    "create_engine"
]
