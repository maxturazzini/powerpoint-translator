from dataclasses import dataclass
from typing import Dict, List, Optional, Any
from docx.text.run import Run
from docx.text.paragraph import Paragraph
from docx.shared import RGBColor
from docx.enum.text import WD_COLOR_INDEX, WD_UNDERLINE
from docx.enum.dml import MSO_THEME_COLOR
import logging

logger = logging.getLogger('word_translator')

@dataclass
class WordRunFormatting:
    """Store all formatting attributes of a Word text run"""
    font_name: Optional[str] = None
    font_size: Optional[float] = None  # Word uses Pt (points) for font size
    bold: Optional[bool] = None  # Can be True, False, or None (inherit)
    italic: Optional[bool] = None
    underline: Optional[Any] = None  # Word has multiple underline types
    color: Optional[Dict[str, Any]] = None  # RGB color info
    highlight_color: Optional[Any] = None  # Word-specific highlighting
    style: Optional[str] = None  # Character style
    all_caps: Optional[bool] = None
    small_caps: Optional[bool] = None
    strike: Optional[bool] = None
    double_strike: Optional[bool] = None
    superscript: Optional[bool] = None
    subscript: Optional[bool] = None
    hidden: Optional[bool] = None
    # Paragraph-level properties
    alignment: Optional[int] = None
    space_before: Optional[float] = None
    space_after: Optional[float] = None
    line_spacing: Optional[float] = None
    keep_together: Optional[bool] = None
    keep_with_next: Optional[bool] = None
    page_break_before: Optional[bool] = None
    widow_control: Optional[bool] = None

    def to_dict(self) -> Dict[str, Any]:
        """Convert formatting to dictionary for storage"""
        return {
            'font_name': self.font_name,
            'font_size': self.font_size,
            'bold': self.bold,
            'italic': self.italic,
            'underline': self.underline,
            'color': self.color,
            'highlight_color': self.highlight_color,
            'style': self.style,
            'all_caps': self.all_caps,
            'small_caps': self.small_caps,
            'strike': self.strike,
            'double_strike': self.double_strike,
            'superscript': self.superscript,
            'subscript': self.subscript,
            'hidden': self.hidden,
            'alignment': self.alignment,
            'space_before': self.space_before,
            'space_after': self.space_after,
            'line_spacing': self.line_spacing,
            'keep_together': self.keep_together,
            'keep_with_next': self.keep_with_next,
            'page_break_before': self.page_break_before,
            'widow_control': self.widow_control
        }

    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> 'WordRunFormatting':
        """Create formatting from dictionary"""
        return cls(**data)

class WordFormattingManager:
    """Manage Word text formatting preservation during translation"""
    
    def __init__(self):
        self.format_maps: Dict[str, List[WordRunFormatting]] = {}
        
    def _get_color_info(self, font) -> Optional[Dict[str, Any]]:
        """Extract color information from Word font"""
        try:
            if hasattr(font, 'color') and font.color:
                color_info = {}
                
                # Get RGB color if available
                if hasattr(font.color, 'rgb') and font.color.rgb:
                    color_info['rgb'] = str(font.color.rgb)
                
                # Get theme color if available
                if hasattr(font.color, 'theme_color') and font.color.theme_color:
                    color_info['theme_color'] = font.color.theme_color
                    
                return color_info if color_info else None
        except Exception as e:
            logger.warning(f"Error getting color info: {str(e)}")
            return None

    def _apply_color_info(self, font, color_info: Dict[str, Any]) -> None:
        """Apply stored color information to Word font"""
        if not color_info:
            return
            
        try:
            if 'rgb' in color_info:
                font.color.rgb = RGBColor.from_string(color_info['rgb'])
            elif 'theme_color' in color_info:
                font.color.theme_color = color_info['theme_color']
        except Exception as e:
            logger.warning(f"Error applying color: {str(e)}")

    def collect_run_formatting(self, run: Run, paragraph: Optional[Paragraph] = None) -> WordRunFormatting:
        """Collect all formatting attributes from a Word text run"""
        try:
            font = run.font
            
            # Collect color information
            color_info = self._get_color_info(font)
            
            # Get paragraph properties if paragraph is provided
            alignment = None
            space_before = None
            space_after = None
            line_spacing = None
            keep_together = None
            keep_with_next = None
            page_break_before = None
            widow_control = None
            
            if paragraph:
                try:
                    alignment = paragraph.alignment
                    space_before = paragraph.paragraph_format.space_before
                    space_after = paragraph.paragraph_format.space_after
                    line_spacing = paragraph.paragraph_format.line_spacing
                    keep_together = paragraph.paragraph_format.keep_together
                    keep_with_next = paragraph.paragraph_format.keep_with_next
                    page_break_before = paragraph.paragraph_format.page_break_before
                    widow_control = paragraph.paragraph_format.widow_control
                except Exception as e:
                    logger.warning(f"Error getting paragraph properties: {str(e)}")
            
            return WordRunFormatting(
                font_name=font.name,
                font_size=font.size.pt if font.size else None,
                bold=font.bold,
                italic=font.italic,
                underline=font.underline,
                color=color_info,
                highlight_color=font.highlight_color,
                style=run.style.name if run.style else None,
                all_caps=font.all_caps,
                small_caps=font.small_caps,
                strike=font.strike,
                double_strike=font.double_strike,
                superscript=font.superscript,
                subscript=font.subscript,
                hidden=font.hidden,
                alignment=alignment,
                space_before=space_before,
                space_after=space_after,
                line_spacing=line_spacing,
                keep_together=keep_together,
                keep_with_next=keep_with_next,
                page_break_before=page_break_before,
                widow_control=widow_control
            )
        except Exception as e:
            logger.error(f"Error collecting run formatting: {str(e)}")
            return WordRunFormatting()

    def store_paragraph_formatting(self, paragraph: Paragraph, paragraph_id: str) -> None:
        """Store formatting for all runs in a paragraph"""
        if paragraph_id not in self.format_maps:
            self.format_maps[paragraph_id] = []
            
        for run in paragraph.runs:
            formatting = self.collect_run_formatting(run, paragraph)
            self.format_maps[paragraph_id].append(formatting)

    def apply_run_formatting(self, run: Run, formatting: WordRunFormatting) -> None:
        """Apply stored formatting to a Word text run"""
        try:
            font = run.font
            
            # Apply font properties
            if formatting.font_name:
                font.name = formatting.font_name
            if formatting.font_size:
                from docx.shared import Pt
                font.size = Pt(formatting.font_size)
            
            # Apply tri-state properties (True, False, or None)
            if formatting.bold is not None:
                font.bold = formatting.bold
            if formatting.italic is not None:
                font.italic = formatting.italic
            if formatting.underline is not None:
                font.underline = formatting.underline
            if formatting.all_caps is not None:
                font.all_caps = formatting.all_caps
            if formatting.small_caps is not None:
                font.small_caps = formatting.small_caps
            if formatting.strike is not None:
                font.strike = formatting.strike
            if formatting.double_strike is not None:
                font.double_strike = formatting.double_strike
            if formatting.superscript is not None:
                font.superscript = formatting.superscript
            if formatting.subscript is not None:
                font.subscript = formatting.subscript
            if formatting.hidden is not None:
                font.hidden = formatting.hidden
            
            # Apply color
            if formatting.color:
                self._apply_color_info(font, formatting.color)
            
            # Apply highlight color
            if formatting.highlight_color is not None:
                font.highlight_color = formatting.highlight_color
                
            # Apply character style
            if formatting.style:
                try:
                    run.style = formatting.style
                except Exception:
                    logger.warning(f"Could not apply style: {formatting.style}")
                    
        except Exception as e:
            logger.warning(f"Error applying run formatting: {str(e)}")

    def apply_paragraph_formatting(self, paragraph: Paragraph, paragraph_id: str, 
                                 format_index: int = 0) -> int:
        """Apply stored formatting to a paragraph"""
        if paragraph_id not in self.format_maps:
            return format_index
            
        stored_formats = self.format_maps[paragraph_id]
        
        # Apply paragraph-level properties from the first format if available
        if format_index < len(stored_formats):
            try:
                formatting = stored_formats[format_index]
                para_format = paragraph.paragraph_format
                
                if formatting.alignment is not None:
                    paragraph.alignment = formatting.alignment
                if formatting.space_before is not None:
                    para_format.space_before = formatting.space_before
                if formatting.space_after is not None:
                    para_format.space_after = formatting.space_after
                if formatting.line_spacing is not None:
                    para_format.line_spacing = formatting.line_spacing
                if formatting.keep_together is not None:
                    para_format.keep_together = formatting.keep_together
                if formatting.keep_with_next is not None:
                    para_format.keep_with_next = formatting.keep_with_next
                if formatting.page_break_before is not None:
                    para_format.page_break_before = formatting.page_break_before
                if formatting.widow_control is not None:
                    para_format.widow_control = formatting.widow_control
            except Exception as e:
                logger.warning(f"Error applying paragraph properties: {str(e)}")
        
        # Apply run-level formatting
        for run in paragraph.runs:
            if format_index < len(stored_formats):
                self.apply_run_formatting(run, stored_formats[format_index])
                format_index += 1
                
        return format_index

    def validate_formatting(self, original_id: str, translated_id: str) -> List[str]:
        """Verify formatting preservation between original and translated paragraphs"""
        warnings = []
        
        if original_id not in self.format_maps:
            warnings.append(f"No formatting data found for original paragraph {original_id}")
            return warnings
            
        if translated_id not in self.format_maps:
            warnings.append(f"No formatting data found for translated paragraph {translated_id}")
            return warnings
            
        original_formats = self.format_maps[original_id]
        translated_formats = self.format_maps[translated_id]
        
        if len(original_formats) != len(translated_formats):
            warnings.append(
                f"Formatting count mismatch: Original={len(original_formats)}, "
                f"Translated={len(translated_formats)}"
            )
            
        for i, (orig_fmt, trans_fmt) in enumerate(
            zip(original_formats, translated_formats)
        ):
            if orig_fmt.to_dict() != trans_fmt.to_dict():
                warnings.append(f"Formatting mismatch at run {i}")
                
        return warnings

    def clear_formatting(self, paragraph_id: str) -> None:
        """Clear stored formatting for a paragraph"""
        if paragraph_id in self.format_maps:
            del self.format_maps[paragraph_id]

    def _should_translate_text(self, text: str) -> bool:
        """Validate if text should be translated"""
        clean_text = text.strip()
        
        # Skip empty, whitespace, single chars
        if not clean_text or text.isspace() or len(clean_text) == 1:
            return False
        
        # Skip uppercase acronyms (2+ characters, all uppercase, not digits)
        if len(clean_text) >= 2 and clean_text.isupper() and not clean_text.isdigit():
            return False
        
        # Skip pure numbers
        if clean_text.isdigit():
            return False
        
        # Skip common non-translatable patterns
        if clean_text in ['&', '...', '•', '→', '←', '↑', '↓']:
            return False
            
        return True