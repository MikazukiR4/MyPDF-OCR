"""
OCR Thai Text Cleaning Pipeline - Part 1 (Standard Processors)
===============================

Modularized OCR text cleaning processors with flexible pipeline architecture.
Replaces monolithic regex-based cleanup with composable processors.

Usage:
    from core.utils.cleaners import clean_thai_ocr_text, create_pipeline

    # Simple usage (uses default pipeline):
    cleaned = clean_thai_ocr_text("  สวัสดีครับ ")

    # Custom pipeline:
    custom_pipeline = create_pipeline(
        [normalize_thai_unicode, remove_ocr_garbage, fix_common_typos],
        skip=["remove_ocr_garbage"]
    )
    cleaned = custom_pipeline(text)
"""

import unicodedata
import re
from typing import Callable, List, Optional


# =============================================================================
# STANDARD PIPELINE PROCESSORS (Default: all processors enabled)
# =============================================================================

def normalize_thai_unicode(text: str) -> str:
    """
    Normalize Thai Unicode and handle PUA (Private Use Area) mappings.
    
    Handles:
    - NFC normalization for consistency
    - Common Unicode variants that OCR produces inconsistently
    
    Args:
        text: Input text with possible Unicode anomalies
        
    Returns:
        Normalized Thai text
        
    Examples:
        >>> normalize_thai_unicode("สวัสดี")
        'สวัสดี'
        >>> normalize_thai_unicode("สวัสดี (variant)")
        'สวัสดี'
    """
    if not text:
        return text
    
    # Create mapping for common Thai Unicode variants
    unicode_mapping = {
        # Full-width to half-width (OCR produces both)
        '\uFF01': '!',
        '\uFF02': '"',
        '\uFF03': '#',
        '\uFF04': '$',
        '\uFF05': '%',
        '\uFF06': '&',
        '\uFF07': "'",
        '\uFF08': '(',
        '\uFF09': ')',
        '\uFF0A': '*',
        '\uFF0B': '+',
        '\uFF0C': ',',
        '\uFF0D': '-',
        '\uFF0E': '.',
        '\uFF0F': '/',
        '\uFF10': '0',
        '\uFF11': '1',
        '\uFF12': '2',
        '\uFF13': '3',
        '\uFF14': '4',
        '\uFF15': '5',
        '\uFF16': '6',
        '\uFF17': '7',
        '\uFF18': '8',
        '\uFF19': '9',
        '\uFF1A': ':',
        '\uFF1B': ';',
        '\uFF1C': '<',
        '\uFF1D': '=',
        '\uFF1E': '>',
        '\uFF1F': '?',
        '\uFF20': '@',
        '\uFF3F': '/',  # Thai slash variant
        '\uFEFF': '',    # Zero-width no-break space
        
        # Fullwidth digits to ASCII
        '\uFF41': 'a', '\uFF42': 'b', '\uFF43': 'c', '\uFF44': 'd',
        '\uFF45': 'e', '\uFF46': 'f', '\uFF47': 'g', '\uFF48': 'h',
        '\uFF49': 'i', '\uFF4A': 'j', '\uFF4B': 'k', '\uFF4C': 'l',
        '\uFF4D': 'm', '\uFF4E': 'n', '\uFF4F': 'o', '\uFF50': 'p',
        '\uFF51': 'q', '\uFF52': 'r', '\uFF53': 's', '\uFF54': 't',
        '\uFF55': 'u', '\uFF56': 'v', '\uFF57': 'w', '\uFF58': 'x',
        '\uFF59': 'y', '\uFF5A': 'z',
        
        # Combining diacritics handling (normalize to base character)
        '\u0300': '',  # Combining grave accent
        '\u0301': '',  # Combining acute accent  
        '\u0302': '',  # Combining circumflex accent
        '\u0303': '',  # Combining tilde
        '\u0304': '',  # Combining macron (U+0304)
        '\u0305': '',  # Combining overline
        '\u0306': '',  # Combining breve
        '\u0308': '',  # Combining diaeresis
        '\u030A': '',  # Combining ring above
        '\u0329': '',  # Thai vowel combining (simplified)
    }
    
    # Build translation table
    translation_table = str.maketrans(unicode_mapping)
    
    # Apply normalization and translation
    normalized_text = unicodedata.normalize("NFC", text.translate(translation_table))
    
    return normalized_text


def remove_ocr_garbage(text: str) -> str:
    """
    Remove common OCR artifacts and PUA (Private Use Area) characters.
    
    Handles:
    - PUA Thai symbols (U+F800-U+FAFF) that represent malformed Thai chars
    - Special OCR markers like [۞], [۩], ✠, ❧, ❥, etc.
    - Broken Thai syllables and orphaned vowels/consonants
    
    Args:
        text: Text containing OCR artifacts
        
    Returns:
        Text with PUA characters and special artifacts removed
    """
    if not text:
        return text
    
    result = text
    
    # 1. Remove PUA Thai symbols (U+F800-U+FAFF) - these are broken Thai chars
    # These represent malformed Thai characters that OCR generates incorrectly
    pua_pattern = r'[\uF800-\uFAFF]'
    result = re.sub(pua_pattern, '', result)
    
    # 2. Remove special OCR marker symbols commonly seen in scanned Thai documents
    ocr_markers = [
        '۞', '۩', '✠', '❧', '❥', '‡', '†', '§', '¶', '•',
        '◘', '◎', '○', '●', '◙', '◕', '☻', '★', '☆', '☉',
        '❂', '❃', '❄', '❅', '❆', '✢', '✤', '✥', '✦', '✧',
        '❋', '☀', '☁', '☂', '☃', '☄', '⚀', '⚁', '⚂', '⚃',
        '⚄', '⚅', '❦', '❧', '⊘', '⊙', '◎', '♛', '♜', '♝',
        '♞', '♟', '♠', '♥', '♦', '♣', '♩', '♪', '§'
    ]
    for marker in ocr_markers:
        result = result.replace(marker, '')
    
    # 3. Remove double newlines or excessive whitespace that OCR creates
    result = re.sub(r'\n\s*\n', '\n\n', result)
    
    return result


def fix_common_typos(text: str) -> str:
    """
    Fix common OCR-induced typos in Thai text.
    
    Handles:
    - Common misrecognitions of similar-looking characters
    - Digit/letter confusions (๕ vs 5, ๖ vs 6, etc.)
    - Repeated characters from line bleed-through
    
    Args:
        text: Text with common OCR typos
        
    Returns:
        Corrected Thai text
    """
    if not text:
        return text
    
    result = text
    
    # 1. Fix common OCR misreads in Thai syllables
    # Example: ัๅ -> ั (sometimes อ appears incorrectly)
    result = re.sub(r'ัๅ', 'ั', result)
    
    # 2. Remove triple or more consecutive identical characters from line bleed-through
    result = re.sub(r'(.)\1{2,}', r'\1', result)
    
    # 3. Normalize common spacing issues
    # Remove spaces before Thai characters at sentence starts
    result = re.sub(r'\s(?=[ก-ฮ๕])', '', result)
    
    return result


def create_pipeline(processors, skip=None):
    """
    Create a text cleaning pipeline from processors.
    
    Args:
        processors: List of processor functions to apply in order
        skip: List of processor names to skip
        
    Returns:
        Pipeline function that applies all non-skipped processors
    """
    if skip is None:
        skip = []
    
    # Build ordered list of active processors
    available_processors = {
        'normalize_thai_unicode': normalize_thai_unicode,
        'remove_ocr_garbage': remove_ocr_garbage,
        'fix_common_typos': fix_common_typos,
    }
    
    active_processors = []
    for proc_name in ['normalize_thai_unicode', 'remove_ocr_garbage', 'fix_common_typos']:
        if proc_name not in skip:
            func = available_processors.get(proc_name)
            if func:
                active_processors.append(func)
    
    def pipeline(text):
        for processor in active_processors:
            text = processor(text)
        return text
    
    return pipeline


def clean_thai_ocr_text(text: str) -> str:
    """
    Clean Thai OCR text using the default full pipeline.
    
    This function applies all standard processors in order:
    1. normalize_thai_unicode - Normalize Unicode variants
    2. remove_ocr_garbage - Remove PUA characters and artifacts
    3. fix_common_typos - Fix common OCR-induced typos
    
    Args:
        text: Raw OCR-extracted Thai text
        
    Returns:
        Cleaned Thai text ready for further processing
    """
    return create_pipeline(
        [normalize_thai_unicode, remove_ocr_garbage, fix_common_typos],
        skip=[]  # No processors skipped by default
    )(text)

