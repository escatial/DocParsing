"""Markdown -> 带脚注的 docx 转换模块。"""
import re
from io import BytesIO
from typing import Dict, List, Tuple, Optional

from docx import Document
from docx.shared import Pt, RGBColor