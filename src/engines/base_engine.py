#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
引擎基類 - 定義統一的模板填充接口

所有引擎（openpyxl, zip）都繼承此基類

作者: David-CB666
版本: v2.1
日期: 2026-06-02
"""

from abc import ABC, abstractmethod
from typing import Dict, List, Any, Optional
from pathlib import Path


class BaseEngine(ABC):
    """
    模板填充引擎基類
    
    定義所有引擎必須實現的接口
    """
    
    def __init__(self):
        """初始化引擎"""
        self.template_path: Optional[str] = None
        self.is_loaded: bool = False
    
    @abstractmethod
    def load_template(self, template_path: str) -> bool:
        """
        載入模板文件
        
        Args:
            template_path: 模板 Excel 路徑
        
        Returns:
            是否成功載入
        """
        pass
    
    @abstractmethod
    def scan_placeholders(self) -> List[str]:
        """
        掃描模板中的佔位符
        
        Returns:
            佔位符列表
        """
        pass
    
    @abstractmethod
    def fill_template(
        self,
        data: Dict[str, Any],
        field_map: Dict[str, str]
    ) -> Any:
        """
        填充單個模板
        
        Args:
            data: 數據字典
            field_map: 字段映射 {佔位符: 字段名}
        
        Returns:
            填充後的對象（具體類型由子類定義）
        """
        pass
    
    @abstractmethod
    def fill_and_export(
        self,
        data_list: List[Dict[str, Any]],
        field_map: Dict[str, str],
        output_path: str
    ) -> str:
        """
        批量填充並導出
        
        Args:
            data_list: 數據列表
            field_map: 字段映射 {佔位符: 字段名}
            output_path: 輸出路徑
        
        Returns:
            輸出文件路徑
        """
        pass
    
    @staticmethod
    def detect_template_type(template_path: str) -> str:
        """
        檢測模板類型（靜態方法）
        
        Args:
            template_path: 模板路徑
        
        Returns:
            "zip" 或 "openpyxl"
        """
        import zipfile
        
        try:
            with zipfile.ZipFile(template_path, 'r') as zf:
                names = zf.namelist()
                
                # 檢查是否包含圖片
                has_images = any(name.startswith('xl/media/') for name in names)
                
                # 檢查是否包含打印設置
                has_printer = any('printerSettings' in name for name in names)
                
                # 檢查是否包含 DrawingML
                has_drawings = any('drawings' in name for name in names)
                
                # 有圖片、打印設置或繪圖 → 使用 ZIP 引擎
                return "zip" if (has_images or has_printer or has_drawings) else "openpyxl"
        
        except Exception:
            return "openpyxl"
    
    def get_template_info(self) -> Dict[str, Any]:
        """
        獲取模板信息
        
        Returns:
            模板信息字典
        """
        if not self.template_path:
            return {}
        
        path = Path(self.template_path)
        
        return {
            "path": str(path),
            "name": path.name,
            "size": path.stat().st_size if path.exists() else 0,
            "type": self.detect_template_type(self.template_path),
            "is_loaded": self.is_loaded
        }
