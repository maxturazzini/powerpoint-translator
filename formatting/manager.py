from dataclasses import dataclass
from typing import Dict, List, Optional, Any
from pptx.text.text import _Run, _Paragraph
from pptx.dml.color import RGBColor, ColorFormat
from pptx.enum.dml import MSO_THEME_COLOR
from pptx.enum.lang import MSO_LANGUAGE_ID
import logging

logger = logging.getLogger('ppt_translator')

@dataclass
class TextRunFormatting:
    """Store all formatting attributes of a text run"""
    font_name: Optional[str] = None
    font_size: Optional[int] = None
    bold: bool = False
    italic: bool = False
    underline: bool = False
    color_format: Optional[Dict[str, Any]] = None  # Store complete color format info
    language_id: Optional[Any] = None  # Store raw language ID value
    spacing: Optional[int] = None  # Character spacing
    alignment: Optional[int] = None  # Paragraph alignment
    level: Optional[int] = None  # Paragraph level

    def to_dict(self) -> Dict[str, Any]:
        """Convert formatting to dictionary for storage"""
        return {
            'font_name': self.font_name,
            'font_size': self.font_size,
            'bold': self.bold,
            'italic': self.italic,
            'underline': self.underline,
            'color_format': self.color_format,
            'language_id': self.language_id,
            'spacing': self.spacing,
            'alignment': self.alignment,
            'level': self.level
        }

    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> 'TextRunFormatting':
        """Create formatting from dictionary"""
        return cls(
            font_name=data.get('font_name'),
            font_size=data.get('font_size'),
            bold=data.get('bold', False),
            italic=data.get('italic', False),
            underline=data.get('underline', False),
            color_format=data.get('color_format'),
            language_id=data.get('language_id'),
            spacing=data.get('spacing'),
            alignment=data.get('alignment'),
            level=data.get('level')
        )

class FormattingManager:
    """Manage text formatting preservation during translation"""
    
    def __init__(self):
        self.format_maps: Dict[str, List[TextRunFormatting]] = {}
        
    def _get_color_format_info(self, color: ColorFormat) -> Optional[Dict[str, Any]]:
        """Extract color information handling different color types"""
        if color is None:
            return None
            
        color_info = {
            'type': color.type,  # Store the color type (RGB, Theme, etc.)
        }
        
        # Handle different color types
        if hasattr(color, 'rgb') and color.rgb is not None:
            color_info['rgb'] = color.rgb
        if hasattr(color, 'theme_color') and color.theme_color is not None:
            # Only store valid theme colors
            if color.theme_color != MSO_THEME_COLOR.NOT_THEME_COLOR:
                color_info['theme_color'] = color.theme_color
        if hasattr(color, 'brightness') and color.brightness is not None:
            color_info['brightness'] = color.brightness
            
        return color_info if any(k != 'type' for k in color_info) else None

    def _apply_color_format(self, font_color: ColorFormat, color_info: Dict[str, Any]) -> None:
        """Apply stored color format information"""
        if not color_info:
            return
            
        color_type = color_info.get('type')
        has_color = False
        
        try:
            # First set the color type (RGB or theme color)
            if color_type == 0 and 'rgb' in color_info:  # MSO_COLOR_TYPE.RGB
                font_color.rgb = RGBColor.from_string(hex(color_info['rgb'])[2:])
                has_color = True
            elif color_type == 1 and 'theme_color' in color_info:  # MSO_COLOR_TYPE.SCHEME
                theme_color = color_info['theme_color']
                # Only set theme color if it's valid
                if theme_color != MSO_THEME_COLOR.NOT_THEME_COLOR:
                    font_color.theme_color = theme_color
                    has_color = True
            
            # Only set brightness if we successfully set a color
            if has_color and 'brightness' in color_info:
                try:
                    font_color.brightness = color_info['brightness']
                except ValueError:
                    # Skip brightness if it can't be set
                    logger.warning("Could not set color brightness")
        except Exception as e:
            # Log any color application errors but continue processing
            logger.warning(f"Error applying color format: {str(e)}")

    def _get_language_id(self, font) -> Optional[int]:
        """Safely get language ID, handling non-standard codes"""
        try:
            # Try to get the language ID directly
            return font.language_id
        except ValueError as e:
            # If it's a non-standard code, default to English
            logger.warning(f"Non-standard language code encountered: {str(e)}")
            return MSO_LANGUAGE_ID.ENGLISH_US
        except Exception as e:
            # For any other error, just skip the language ID
            logger.warning(f"Error getting language ID: {str(e)}")
            return None

    def collect_run_formatting(self, run: _Run, paragraph: Optional[_Paragraph] = None) -> TextRunFormatting:
        """Collect all formatting attributes from a text run"""
        font = run.font
        
        # Extract color information safely
        color_info = self._get_color_format_info(font.color)
        
        # Get language ID safely
        language_id = self._get_language_id(font)
        
        # Get paragraph properties if paragraph is provided
        alignment = None
        level = None
        if paragraph:
            try:
                alignment = paragraph.alignment
                level = paragraph.level
            except Exception as e:
                logger.warning(f"Error getting paragraph properties: {str(e)}")
            
        return TextRunFormatting(
            font_name=font.name,
            font_size=font.size,
            bold=font.bold,
            italic=font.italic,
            underline=font.underline,
            color_format=color_info,
            language_id=language_id,
            spacing=getattr(font, 'spacing', None),
            alignment=alignment,
            level=level
        )

    def store_paragraph_formatting(self, paragraph: _Paragraph, shape_id: str) -> None:
        """Store formatting for all runs in a paragraph"""
        if shape_id not in self.format_maps:
            self.format_maps[shape_id] = []
            
        for run in paragraph.runs:
            formatting = self.collect_run_formatting(run, paragraph)
            self.format_maps[shape_id].append(formatting)

    def apply_run_formatting(self, run: _Run, formatting: TextRunFormatting) -> None:
        """Apply stored formatting to a new text run"""
        try:
            font = run.font
            
            if formatting.font_name:
                font.name = formatting.font_name
            if formatting.font_size:
                font.size = formatting.font_size
            
            font.bold = formatting.bold
            font.italic = formatting.italic
            font.underline = formatting.underline
            
            if formatting.color_format:
                self._apply_color_format(font.color, formatting.color_format)
                
            # Apply language ID safely
            if formatting.language_id is not None:
                try:
                    font.language_id = formatting.language_id
                except ValueError:
                    # If setting fails, try using English
                    try:
                        font.language_id = MSO_LANGUAGE_ID.ENGLISH_US
                    except:
                        logger.warning("Could not set language ID")
                    
            if formatting.spacing is not None and hasattr(font, 'spacing'):
                font.spacing = formatting.spacing
        except Exception as e:
            logger.warning(f"Error applying run formatting: {str(e)}")

    def apply_paragraph_formatting(self, paragraph: _Paragraph, shape_id: str, 
                                 format_index: int = 0) -> int:
        """Apply stored formatting to a new paragraph"""
        if shape_id not in self.format_maps:
            return format_index
            
        stored_formats = self.format_maps[shape_id]
        
        # First apply paragraph-level properties from the first format
        if format_index < len(stored_formats):
            try:
                if stored_formats[format_index].alignment is not None:
                    paragraph.alignment = stored_formats[format_index].alignment
                if stored_formats[format_index].level is not None:
                    paragraph.level = stored_formats[format_index].level
            except Exception as e:
                logger.warning(f"Error applying paragraph properties: {str(e)}")
        
        # Then apply run-level formatting
        for run in paragraph.runs:
            if format_index < len(stored_formats):
                self.apply_run_formatting(run, stored_formats[format_index])
                format_index += 1
                
        return format_index

    def validate_formatting(self, original_shape_id: str, translated_shape_id: str) -> List[str]:
        """Verify formatting preservation between original and translated shapes"""
        warnings = []
        
        if original_shape_id not in self.format_maps:
            warnings.append(f"No formatting data found for original shape {original_shape_id}")
            return warnings
            
        if translated_shape_id not in self.format_maps:
            warnings.append(f"No formatting data found for translated shape {translated_shape_id}")
            return warnings
            
        original_formats = self.format_maps[original_shape_id]
        translated_formats = self.format_maps[translated_shape_id]
        
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

    def clear_formatting(self, shape_id: str) -> None:
        """Clear stored formatting for a shape"""
        if shape_id in self.format_maps:
            del self.format_maps[shape_id]
